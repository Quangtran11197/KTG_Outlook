#![cfg_attr(all(windows, not(debug_assertions)), windows_subsystem = "windows")]
//! Desktop window and private loopback worker. One STA worker handles each job.
#[cfg(windows)]
mod desktop;
use std::{
    collections::BTreeMap,
    fs,
    io::{self, Read, Write},
    net::{TcpListener, TcpStream},
    path::PathBuf,
    process::{Child, Command, Stdio},
    sync::{
        atomic::{AtomicBool, Ordering},
        Arc,
    },
    thread,
    time::{Duration, SystemTime, UNIX_EPOCH},
};
const LIMIT: usize = 64 * 1024 * 1024;
const HTML: &str = include_str!("../web/index.html");
const JS: &str = include_str!("../web/app.js");
const CSS: &str = include_str!("../web/style.css");
const CORE: &str = include_str!("../scripts/core.ps1");
const BRIDGE: &str = include_str!("../scripts/bridge.ps1");
const BULK: &str = include_str!("../scripts/bulk.ps1");
const MAPI_DLL: &[u8] = include_bytes!("../native/bin/ktg-mapi-x64.dll");

#[derive(Debug)]
struct Request {
    method: String,
    target: String,
    headers: BTreeMap<String, String>,
    body: Vec<u8>,
}
fn read_request(reader: &mut impl Read) -> io::Result<Request> {
    let bad = || io::Error::new(io::ErrorKind::InvalidData, "Invalid HTTP request");
    let mut header = Vec::new();
    let mut byte = [0];
    loop {
        reader.read_exact(&mut byte)?;
        header.push(byte[0]);
        if header.ends_with(b"\r\n\r\n") {
            break;
        }
        if header.len() > 16384 {
            return Err(bad());
        }
    }
    let text = std::str::from_utf8(&header).map_err(|_| bad())?;
    let mut lines = text.split("\r\n");
    let first: Vec<_> = lines.next().ok_or_else(bad)?.split_whitespace().collect();
    if first.len() != 3 || first[2] != "HTTP/1.1" {
        return Err(bad());
    }
    let mut headers = BTreeMap::new();
    for line in lines.filter(|s| !s.is_empty()) {
        let (key, value) = line.split_once(':').ok_or_else(bad)?;
        if headers
            .insert(key.trim().to_ascii_lowercase(), value.trim().to_string())
            .is_some()
        {
            return Err(bad());
        }
    }
    if headers.contains_key("transfer-encoding") {
        return Err(bad());
    }
    let length = headers
        .get("content-length")
        .map(|s| s.parse::<usize>())
        .transpose()
        .map_err(|_| bad())?
        .unwrap_or(0);
    if length > LIMIT {
        return Err(bad());
    }
    let mut body = vec![0; length];
    reader.read_exact(&mut body)?;
    Ok(Request {
        method: first[0].into(),
        target: first[1].into(),
        headers,
        body,
    })
}
fn respond(stream: &mut TcpStream, status: &str, mime: &str, body: &[u8]) -> io::Result<()> {
    write!(stream,"HTTP/1.1 {status}\r\nContent-Type: {mime}\r\nContent-Length: {}\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nReferrer-Policy: no-referrer\r\nContent-Security-Policy: default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'\r\nConnection: close\r\n\r\n",body.len())?;
    stream.write_all(body)
}
fn token() -> io::Result<String> {
    let mut bytes = [0u8; 32];
    #[cfg(windows)]
    unsafe {
        #[link(name = "bcrypt")]
        extern "system" {
            fn BCryptGenRandom(
                algorithm: *mut std::ffi::c_void,
                buffer: *mut u8,
                length: u32,
                flags: u32,
            ) -> i32;
        }
        if BCryptGenRandom(std::ptr::null_mut(), bytes.as_mut_ptr(), 32, 2) < 0 {
            return Err(io::Error::other("Random generator failed"));
        }
    }
    #[cfg(not(windows))]
    {
        fs::File::open("/dev/urandom")?.read_exact(&mut bytes)?;
    }
    Ok(bytes.iter().map(|b| format!("{b:02x}")).collect())
}
fn hidden(command: &mut Command) {
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        command.creation_flags(0x08000000);
    }
    #[cfg(not(windows))]
    {
        let _ = command;
    }
}
fn powershell() -> PathBuf {
    PathBuf::from(std::env::var_os("SystemRoot").unwrap_or_else(|| "C:\\Windows".into()))
        .join("System32/WindowsPowerShell/v1.0/powershell.exe")
}
fn pick_pst(mode: &str) -> io::Result<Vec<u8>> {
    let class = match mode {
        "existing" => "OpenFileDialog",
        "new" => "SaveFileDialog",
        _ => return Ok(br#"{"error":"Invalid PST mode"}"#.to_vec()),
    };
    let script = format!(
        r#"$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
Add-Type -AssemblyName System.Windows.Forms
$dialog=New-Object System.Windows.Forms.{class}
$dialog.Filter='Outlook Data Files (*.pst)|*.pst'
$dialog.DefaultExt='pst'
$dialog.AddExtension=$true
$dialog.RestoreDirectory=$true
$dialog.Title='Chọn file PST'
if('{mode}' -eq 'new'){{$dialog.OverwritePrompt=$false;$dialog.FileName='Archive.pst'}}
try{{if($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK){{$dialog.FileName|ConvertTo-Json -Compress}}else{{'null'}}}}finally{{$dialog.Dispose()}}
"#
    );
    let mut command = Command::new(powershell());
    command.args(["-NoProfile", "-Sta", "-Command", &script]);
    hidden(&mut command);
    let output = command.output()?;
    if !output.status.success() {
        return Ok(br#"{"error":"Cannot open PST picker; enter the path manually"}"#.to_vec());
    }
    let text = String::from_utf8_lossy(&output.stdout);
    Ok(text.trim().as_bytes().to_vec())
}
struct Job {
    path: PathBuf,
    child: Child,
}
#[derive(Default)]
struct ServerControl {
    shutdown: AtomicBool,
    cancel: AtomicBool,
    finished: AtomicBool,
}
fn read_state(job: &mut Job) -> Vec<u8> {
    let state=fs::read(job.path.join("state.json")).unwrap_or_else(|_|br#"{"status":"running","processed":0,"total":0,"copied":0,"errors":0,"batches":0,"resultBatches":0,"elapsedSeconds":0}"#.to_vec());
    if matches!(job.child.try_wait(), Ok(Some(_)))
        && String::from_utf8_lossy(&state).contains("\"status\":\"running\"")
    {
        return br#"{"status":"error","note":"Worker exited unexpectedly. Check worker-error.log and report.csv; do not blindly retry.","errors":1,"batches":0,"resultBatches":0}"#.to_vec();
    }
    state
}
fn batch_name(target: &str) -> Option<String> {
    let query = target.strip_prefix("/api/batch?")?;
    let mut kind = None;
    let mut index = None;
    for pair in query.split('&') {
        let (k, v) = pair.split_once('=')?;
        match k {
            "kind" if kind.is_none() => kind = Some(v),
            "index" if index.is_none() => index = Some(v.parse::<u32>().ok()?),
            _ => return None,
        }
    }
    let kind = kind?;
    if !matches!(kind, "messages" | "results") {
        return None;
    }
    Some(format!("{kind}-{}.json", index?))
}
fn run() -> io::Result<()> {
    let args: Vec<_> = std::env::args().collect();
    let root = std::env::current_exe()?
        .parent()
        .ok_or_else(|| io::Error::other("Missing application directory"))?
        .to_path_buf();
    let data = if let Some(i) = args.iter().position(|a| a == "--data-dir") {
        PathBuf::from(
            args.get(i + 1)
                .ok_or_else(|| io::Error::other("Missing --data-dir value"))?,
        )
    } else {
        std::env::var_os("APPDATA")
            .map(PathBuf::from)
            .map(|p| p.join("KTG/OutlookMailManager"))
            .unwrap_or_else(|| root.join("data"))
    };
    let key = token()?;
    let session = data.join(format!("session-{}", &key[..12]));
    fs::create_dir_all(&session)?;
    for (name, content) in [
        ("core.ps1", CORE),
        ("bridge.ps1", BRIDGE),
        ("bulk.ps1", BULK),
    ] {
        fs::write(
            session.join(name),
            format!("\u{feff}{}", content.trim_start_matches('\u{feff}')),
        )?;
    }
    fs::write(session.join("ktg-mapi-x64.dll"), MAPI_DLL)?;
    let listener = TcpListener::bind("127.0.0.1:0")?;
    let address = listener.local_addr()?.to_string();
    let origin = format!("http://{address}");
    let url = format!("{origin}/?token={key}");
    println!("{url}");
    println!("Data: {}", session.display());
    if let Some(i) = args.iter().position(|a| a == "--ready-file") {
        let path = args
            .get(i + 1)
            .ok_or_else(|| io::Error::other("Missing --ready-file value"))?;
        fs::write(
            path,
            format!("{{\"url\":\"{url}\",\"token\":\"{key}\",\"origin\":\"{origin}\"}}"),
        )?;
    }
    let control = Arc::new(ServerControl::default());
    if args.iter().any(|a| a == "--headless") {
        return serve(listener, session, key, control);
    }
    #[cfg(windows)]
    {
        let worker_control = control.clone();
        let server = thread::spawn(move || {
            let result = serve(listener, session, key, worker_control.clone());
            worker_control.finished.store(true, Ordering::SeqCst);
            result
        });
        let gui_result = desktop::show(&url, &origin, &data, control.clone());
        control.shutdown.store(true, Ordering::SeqCst);
        let server_result = server
            .join()
            .map_err(|_| io::Error::other("Local worker stopped unexpectedly"))?;
        gui_result?;
        server_result
    }
    #[cfg(not(windows))]
    {
        serve(listener, session, key, control)
    }
}
fn serve(
    listener: TcpListener,
    session: PathBuf,
    key: String,
    control: Arc<ServerControl>,
) -> io::Result<()> {
    let address = listener.local_addr()?.to_string();
    let origin = format!("http://{address}");
    listener.set_nonblocking(true)?;
    let mut job: Option<Job> = None;
    let mut sequence = 0u64;
    let result = (|| {
        loop {
            let active = if let Some(j) = job.as_mut() {
                j.child.try_wait()?.is_none()
            } else {
                false
            };
            if control.cancel.swap(false, Ordering::SeqCst)
                || control.shutdown.load(Ordering::SeqCst)
            {
                if let Some(j) = &job {
                    if active {
                        fs::write(j.path.join("cancel"), b"cancel")?;
                    }
                }
            }
            if control.shutdown.load(Ordering::SeqCst) && !active {
                break;
            }
            let mut stream = match listener.accept() {
                Ok((stream, _)) => stream,
                Err(e) if e.kind() == io::ErrorKind::WouldBlock => {
                    thread::sleep(Duration::from_millis(30));
                    continue;
                }
                Err(e) => return Err(e),
            };
            // Windows accepts inherit the listener's nonblocking setting.
            // Request bodies can arrive after their headers, so use blocking I/O here.
            stream.set_nonblocking(false)?;
            stream.set_read_timeout(Some(Duration::from_secs(10)))?;
            stream.set_write_timeout(Some(Duration::from_secs(10)))?;
            let request = match read_request(&mut stream) {
                Ok(r) => r,
                Err(_) => {
                    let _ = respond(
                        &mut stream,
                        "400 Bad Request",
                        "text/plain",
                        b"Invalid or oversized request",
                    );
                    continue;
                }
            };
            let host_ok = request.headers.get("host") == Some(&address);
            let origin_ok = request
                .headers
                .get("origin")
                .map(|s| s == &origin)
                .unwrap_or(true);
            let api = request.target.starts_with("/api/");
            if !host_ok || !origin_ok || (api && request.headers.get("x-ktg-token") != Some(&key)) {
                let _ = respond(&mut stream, "403 Forbidden", "text/plain", b"Forbidden");
                continue;
            }
            let mut code = "200 OK";
            let mut mime = "application/json; charset=utf-8";
            let body = match (request.method.as_str(), request.target.as_str()) {
                ("GET", path) if path == "/" || path.starts_with("/?token=") => {
                    HTML.as_bytes().to_vec()
                }
                ("GET", "/app.js") => JS.as_bytes().to_vec(),
                ("GET", "/style.css") => CSS.as_bytes().to_vec(),
                ("GET", "/api/state") => {
                    if let Some(j) = job.as_mut() {
                        read_state(j)
                    } else {
                        br#"{"status":"idle","processed":0,"total":0,"copied":0,"errors":0,"batches":0,"resultBatches":0,"elapsedSeconds":0}"#.to_vec()
                    }
                }
                ("GET", path) if batch_name(path).is_some() => job
                    .as_ref()
                    .and_then(|j| fs::read(j.path.join(batch_name(path).unwrap())).ok())
                    .unwrap_or_else(|| b"[]".to_vec()),
                ("GET", "/api/report") => job
                    .as_ref()
                    .and_then(|j| fs::read(j.path.join("report.csv")).ok())
                    .unwrap_or_default(),
                ("POST", "/api/pickpst?mode=existing") if !active => pick_pst("existing")?,
                ("POST", "/api/pickpst?mode=new") if !active => pick_pst("new")?,
                ("POST", "/api/start") if !active && !control.shutdown.load(Ordering::SeqCst) => {
                    sequence += 1;
                    let time = SystemTime::now()
                        .duration_since(UNIX_EPOCH)
                        .unwrap_or_default()
                        .as_millis();
                    let path = session.join(format!("job-{time}-{sequence}"));
                    fs::create_dir(&path)?;
                    fs::write(path.join("request.json"), &request.body)?;
                    let err = fs::File::create(path.join("worker-error.log"))?;
                    let mut command = Command::new(powershell());
                    command
                        .args([
                            "-NoProfile",
                            "-NonInteractive",
                            "-Sta",
                            "-ExecutionPolicy",
                            "Bypass",
                            "-File",
                        ])
                        .arg(session.join("bridge.ps1"))
                        .arg("-InputFile")
                        .arg(path.join("request.json"));
                    command.env(
                        "KTG_LOG_FILE",
                        session.parent().unwrap_or(&session).join("operations.log"),
                    );
                    command
                        .stdin(Stdio::null())
                        .stdout(Stdio::null())
                        .stderr(Stdio::from(err));
                    hidden(&mut command);
                    match command.spawn() {
                        Ok(child) => {
                            job = Some(Job { path, child });
                            br#"{"started":true}"#.to_vec()
                        }
                        Err(_) => br#"{"error":"Cannot start Windows PowerShell"}"#.to_vec(),
                    }
                }
                ("POST", "/api/start") => {
                    br#"{"error":"A job is running or the application is closing"}"#.to_vec()
                }
                ("POST", "/api/pause") if active => {
                    if let Some(j) = &job {
                        let marker = j.path.join("pause");
                        if marker.exists() {
                            fs::remove_file(marker)?;
                        } else {
                            fs::write(marker, b"pause")?;
                        }
                    }
                    br#"{"pauseToggled":true}"#.to_vec()
                }
                ("POST", "/api/cancel") => {
                    if let Some(j) = &job {
                        fs::write(j.path.join("cancel"), b"cancel")?;
                    }
                    br#"{"cancelRequested":true}"#.to_vec()
                }
                ("POST", "/api/quit") if !active => {
                    respond(
                        &mut stream,
                        "200 OK",
                        "application/json",
                        br#"{"closed":true}"#,
                    )?;
                    break;
                }
                _ => br#"{"error":"Unsupported request"}"#.to_vec(),
            };
            match request.target.split('?').next().unwrap_or("") {
                "/" => mime = "text/html; charset=utf-8",
                "/app.js" => mime = "text/javascript; charset=utf-8",
                "/style.css" => mime = "text/css; charset=utf-8",
                "/api/report" => mime = "text/csv; charset=utf-8",
                _ => {}
            }
            if body.starts_with(br#"{"error":"#) {
                code = "400 Bad Request";
            }
            let _ = respond(&mut stream, code, mime, &body);
        }
        Ok(())
    })();
    // A close request or a server error must never forcibly terminate a mail copy.
    if let Some(j) = job.as_mut() {
        if j.child.try_wait()?.is_none() {
            let _ = fs::write(j.path.join("cancel"), b"cancel");
            while j.child.try_wait()?.is_none() {
                thread::sleep(Duration::from_millis(100));
            }
        }
    }
    control.finished.store(true, Ordering::SeqCst);
    result
}
fn main() {
    if let Err(e) = run() {
        eprintln!("KTG: {e}");
        #[cfg(windows)]
        if !std::env::args().any(|a| a == "--headless") {
            desktop::show_error(&e.to_string());
        }
        std::process::exit(1)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn body_roundtrip() {
        let r = read_request(
            &mut &b"POST /api/start HTTP/1.1\r\nHost: 127.0.0.1:10\r\nContent-Length: 2\r\n\r\n{}"
                [..],
        )
        .unwrap();
        assert_eq!(r.body, b"{}");
    }
    #[test]
    fn reject_smuggling() {
        for s in [
            "POST / HTTP/1.1\r\nContent-Length: 0\r\nContent-Length: 1\r\n\r\n",
            "POST / HTTP/1.1\r\nTransfer-Encoding: chunked\r\n\r\n",
            "POST / HTTP/1.1\r\nContent-Length: 999999999\r\n\r\n",
        ] {
            assert!(read_request(&mut s.as_bytes()).is_err());
        }
    }
    #[test]
    fn reject_traversal() {
        assert!(batch_name("/api/batch?kind=../../request&index=0").is_none());
        assert!(batch_name("/api/batch?kind=messages&index=../0").is_none());
        assert_eq!(
            batch_name("/api/batch?kind=messages&index=40"),
            Some("messages-40.json".into())
        );
    }
    #[test]
    fn token_is_unique() {
        assert_ne!(token().unwrap(), token().unwrap());
        assert_eq!(token().unwrap().len(), 64);
    }
    #[test]
    fn idle_close_stops_server() {
        let listener = TcpListener::bind("127.0.0.1:0").unwrap();
        let control = Arc::new(ServerControl::default());
        control.shutdown.store(true, Ordering::SeqCst);
        let folder = std::env::temp_dir().join(format!("ktg-idle-{}", token().unwrap()));
        serve(listener, folder, token().unwrap(), control.clone()).unwrap();
        assert!(control.finished.load(Ordering::SeqCst));
    }
    #[cfg(windows)]
    #[test]
    fn close_cancels_and_waits_for_worker() {
        let folder = std::env::temp_dir().join(format!("ktg-close-{}", token().unwrap()));
        fs::create_dir_all(&folder).unwrap();
        // This worker is a fixture only. It never creates an Outlook COM object.
        fs::write(
            folder.join("bridge.ps1"),
            br#"param([string]$InputFile)
$dir=[IO.Path]::GetDirectoryName($InputFile)
[IO.File]::WriteAllText((Join-Path $dir 'started'),'yes')
while(-not [IO.File]::Exists((Join-Path $dir 'cancel'))){Start-Sleep -Milliseconds 20}
while(-not [IO.File]::Exists((Join-Path $dir 'release'))){Start-Sleep -Milliseconds 20}
[IO.File]::WriteAllText((Join-Path $dir 'worker-finished'),'yes')
"#,
        )
        .unwrap();

        let listener = TcpListener::bind("127.0.0.1:0").unwrap();
        let address = listener.local_addr().unwrap();
        let control = Arc::new(ServerControl::default());
        let worker_control = control.clone();
        let session = folder.clone();
        let key = token().unwrap();
        let worker_key = key.clone();
        let server = thread::spawn(move || serve(listener, session, worker_key, worker_control));
        let mut stream = TcpStream::connect(address).unwrap();
        stream
            .set_read_timeout(Some(Duration::from_secs(10)))
            .unwrap();
        write!(stream, "POST /api/start HTTP/1.1\r\nHost: {address}\r\nx-ktg-token: {key}\r\nContent-Length: 2\r\n\r\n{{}}").unwrap();
        let mut response = String::new();
        stream.read_to_string(&mut response).unwrap();
        let deadline = std::time::Instant::now() + Duration::from_secs(10);
        let mut job_folder = None;
        while std::time::Instant::now() < deadline {
            job_folder = fs::read_dir(&folder)
                .unwrap()
                .filter_map(Result::ok)
                .find(|e| e.path().is_dir())
                .map(|e| e.path());
            if job_folder
                .as_ref()
                .is_some_and(|p| p.join("started").exists())
            {
                break;
            }
            thread::sleep(Duration::from_millis(30));
        }
        control.shutdown.store(true, Ordering::SeqCst);
        let deadline = std::time::Instant::now() + Duration::from_secs(10);
        while std::time::Instant::now() < deadline
            && !job_folder
                .as_ref()
                .is_some_and(|p| p.join("cancel").exists())
        {
            thread::sleep(Duration::from_millis(30));
        }
        let cancelled = job_folder
            .as_ref()
            .is_some_and(|p| p.join("cancel").exists());
        let waited = !control.finished.load(Ordering::SeqCst);
        // Always release the synthetic worker before asserting, including on failure.
        if let Some(p) = &job_folder {
            fs::write(p.join("release"), b"yes").unwrap();
        }
        let result = server.join().unwrap();
        assert!(response.contains("\"started\":true"));
        assert!(cancelled, "closing must request cooperative cancellation");
        assert!(waited, "server must stay alive until the worker finishes");
        assert!(job_folder.unwrap().join("worker-finished").exists());
        result.unwrap();
        assert!(control.finished.load(Ordering::SeqCst));
    }
}
