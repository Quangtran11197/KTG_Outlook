use crate::ServerControl;
use std::{
    io,
    path::Path,
    sync::{atomic::Ordering, Arc},
    time::{Duration, Instant},
};
use tao::{
    dpi::LogicalSize,
    event::{Event, WindowEvent},
    event_loop::{ControlFlow, EventLoop},
    platform::run_return::EventLoopExtRunReturn,
    window::WindowBuilder,
};
use wry::{NewWindowResponse, WebContext, WebViewBuilder, WebViewBuilderExtWindows};

pub fn show(url: &str, origin: &str, data: &Path, control: Arc<ServerControl>) -> io::Result<()> {
    let mut events = EventLoop::new();
    let window = WindowBuilder::new()
        .with_title("KTG Outlook Mail Manager")
        .with_inner_size(LogicalSize::new(1280.0, 880.0))
        .with_min_inner_size(LogicalSize::new(800.0, 600.0))
        .build(&events)
        .map_err(|e| io::Error::other(e.to_string()))?;
    // A dedicated profile avoids sharing cookies/settings with the user's browser.
    let mut context = WebContext::new(Some(data.join("webview-profile")));
    let allowed = format!("{origin}/");
    let webview = WebViewBuilder::new_with_web_context(&mut context)
        .with_url(url)
        .with_devtools(false)
        .with_browser_accelerator_keys(false)
        .with_default_context_menus(false)
        .with_new_window_req_handler(|_, _| NewWindowResponse::Deny)
        .with_navigation_handler(move |url| url.starts_with(&allowed))
        .build(&window)
        .map_err(|e| {
            io::Error::other(format!(
                "Không mở được WebView2. Hãy kiểm tra Microsoft Edge WebView2 Runtime. {e}"
            ))
        })?;
    let mut closing = false;
    events.run_return(|event, _, flow| {
        *flow = ControlFlow::WaitUntil(Instant::now() + Duration::from_millis(150));
        match event {
            Event::WindowEvent {
                event: WindowEvent::CloseRequested,
                ..
            } => {
                // Keep the window alive until the worker acknowledges completion/cancellation.
                control.shutdown.store(true, Ordering::SeqCst);
                control.cancel.store(true, Ordering::SeqCst);
                if !closing {
                    closing = true;
                    window.set_title("KTG Outlook Mail Manager — Đang dừng tác vụ…");
                    let _ = webview
                        .evaluate_script("window.ktgRequestClose && window.ktgRequestClose()");
                }
            }
            Event::MainEventsCleared if control.finished.load(Ordering::SeqCst) => {
                *flow = ControlFlow::Exit;
            }
            _ => {}
        }
    });
    // Explicitly drop the WebView before its context and native parent.
    drop(webview);
    drop(context);
    drop(window);
    Ok(())
}

pub fn show_error(message: &str) {
    #[link(name = "user32")]
    extern "system" {
        fn MessageBoxW(
            parent: *mut std::ffi::c_void,
            text: *const u16,
            title: *const u16,
            flags: u32,
        ) -> i32;
    }
    let text: Vec<_> = message.encode_utf16().chain(Some(0)).collect();
    let title: Vec<_> = "KTG Outlook Mail Manager"
        .encode_utf16()
        .chain(Some(0))
        .collect();
    unsafe {
        MessageBoxW(std::ptr::null_mut(), text.as_ptr(), title.as_ptr(), 0x10);
    }
}
