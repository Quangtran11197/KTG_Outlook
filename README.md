# KTG Outlook Mail Manager — dự án Rust dựng lại

Bản 0.5.23, ngày 17/11/2026. Đây là dự án mới dựa trên luồng xử lý và giao diện đọc được từ EXE người dùng cung cấp, không phải mã Rust gốc được khôi phục nguyên vẹn.

## Khởi chạy

1. Giải nén toàn bộ dự án vào thư mục có quyền ghi, ngoài OneDrive.
2. Mở Outlook **Classic**, chọn đúng profile rồi chọn **Send/Receive → Work Offline**.
   Chỉ các mailbox Exchange có IsCachedExchange=true trong profile đang mở được quét; bỏ qua PST và hộp thư chỉ online. Công cụ không tự đọc các file OST rời hoặc cache của profile chưa mở.
3. Mở **KTG-Outlook-Mail-Manager-Desktop-0.5.23.exe** hoặc **Start.cmd**. Giao diện mở trong cửa sổ ứng dụng Windows riêng.
4. Quét thư, chọn vài thư trước, nhập đường dẫn PST mới hoặc có sẵn. Thư mục chứa PST phải tồn tại.
5. Sao chép thử một nhóm nhỏ. Mở PST trong Outlook, kiểm tra nội dung và tệp đính kèm rồi mới chạy lô lớn.
6. Tải báo cáo CSV sau mỗi lần xuất. Bấm **Đóng ứng dụng** hoặc nút X để thoát.

Gói này có **KTG-Outlook-Mail-Manager-Desktop-0.5.23.exe bản Rust Windows x64**. Giao diện và script Outlook được nhúng sẵn. Cửa sổ dùng Microsoft Edge WebView2 Runtime đã có trên máy, không có thanh địa chỉ/tab trình duyệt. WebView2Loader được liên kết tĩnh, không cần chép DLL này bên cạnh EXE.

Start.cmd chỉ mở ứng dụng desktop; nếu thiếu EXE, cần giải nén đủ gói hoặc build lại. Không tự chuyển sang trình duyệt. scripts/server.ps1 và cờ --headless được giữ riêng cho kiểm thử/phát triển. EXE sử dụng Microsoft Visual C++ runtime; runtime hiện có trên máy đã cài công cụ build.

Không cần quyền Administrator. Outlook và công cụ cần chạy cùng tài khoản/quyền. Không tự mở Outlook, đổi profile, đổi Work Offline hoặc kết nối Exchange.

## Bảo vệ thư gốc và giới hạn

- Xuất PST có tùy chọn **Tự động xóa thư nguồn** bật sẵn cho cả PST có sẵn và PST mới. Xác nhận một lần trước khi chạy; sau khi lưu thành công không hỏi lại. Bỏ chọn để chỉ sao chép và giữ nguồn.
- Bước sao chép gọi Copy trên thư nguồn, rồi Move **bản sao vừa tạo** sang PST. Chỉ sau khi kiểm tra bản PST mới bắt đầu tự xóa nguồn.
- Chọn **Xóa các email đã chọn** rồi bấm **Thực hiện lựa chọn** sẽ chạy ngay, không hộp xác nhận và không nhập XOA N. Chỉ xử lý các thư đã chọn, chuyển vào Deleted Items rồi mở lại, xóa và kiểm tra kết quả cục bộ. Khi tắt tự xóa và dùng PST có sẵn, sau khi sao chép thành công có nút **Xác nhận đã lưu PST và xóa N email nguồn**; không cần chọn lại các thư thành công.
- Tự xóa sau xuất chỉ áp dụng cho các bản đã được kiểm tra trong phiên xuất đó. Bản native tái dùng biên nhận PST đã được xác minh trong bước copy; nguồn chỉ đủ điều kiện khi có ID, Search Key, timestamp và metadata trong biên nhận. Nếu copy/PST chưa xác minh hoặc thiếu biên nhận, không xóa nguồn.
- Trước khi xóa, kiểm tra EntryID, StoreID, thư mục nguồn, tiêu đề và Internet Message ID nếu có. Chỉ chấp nhận mailbox Exchange cached. Nếu thư đã thay đổi, yêu cầu quét lại.
- Thao tác xóa chuyển thư vào Deleted Items, giải phóng đối tượng cũ và mở lại đúng EntryID trong Deleted Items trước khi gọi Delete. Sau đó chỉ ghi thành công khi GetItemFromID trả lỗi MAPI không còn mục; lệnh không báo lỗi nhưng thư vẫn tồn tại sẽ được ghi là thất bại. Báo cáo ghi ý định trước khi chuyển, ghi ID sau khi chuyển và kết quả. Chính sách lưu giữ Exchange có thể vẫn áp dụng; trạng thái hiển thị là **Đã yêu cầu xóa**, không khẳng định dữ liệu đã bị xóa khỏi máy chủ.
- Tự dọn của bản native dùng DeleteMessages với DELETE_HARD_DELETE trực tiếp trên các ID nguồn sau khi PST đã xác minh, không chuyển chúng vào Deleted Items trước. Đây là xóa vĩnh viễn trong Outlook; không tự thử lại khi kết quả một phần/không xác định. Chính sách Exchange/Recoverable Items vẫn có thể áp dụng.
- Đọc DownloadState trước khi sao chép; chỉ chấp nhận giá trị 1 (đã tải đầy đủ). Chế độ online, thiếu cache hoặc không đọc được trạng thái sẽ bị từ chối.
- Kiểm tra loại thư, EntryID nguồn, StoreID/thư mục đích, tiêu đề, số lượng tệp đính kèm, dung lượng dương và Internet Message ID nếu có.
- Đây là kiểm tra thuộc tính, **chưa phải so sánh byte nội dung hoặc hash từng tệp đính kèm**. Mở thư và tệp trong PST để xác nhận độc lập trước mọi quyết định xóa thủ công.
- Copy có thể tạo bản sao tạm trong mailbox/OST. Nếu mailbox/OST đã đầy, phương án này có thể thất bại. Khi tắt tự xóa, xuất PST không giải phóng dung lượng mailbox vì giữ lại thư gốc.
- Nếu chuyển bản sao PST hoặc kiểm tra bị lỗi, giữ bản sao ở nơi Outlook đã tạo và ghi ID vào báo cáo; không tự xóa hay tự thử lại. Thư nguồn tương ứng được giữ lại nếu sao chép hoặc kiểm tra bị lỗi.
- Khi dừng, chờ thao tác Outlook hiện tại kết thúc. Không cưỡng ép tắt Outlook. Tác vụ có thể chờ lâu nếu COM bị treo.
- Nút **Dừng tác vụ** dùng cho quét, xuất PST và xóa. Khi quét, kiểm tra yêu cầu dừng cả bên trong lô dữ liệu, giữ lại danh sách thư đã đọc.
- Khi bấm nút X lúc tác vụ đang chạy, cửa sổ chuyển sang trạng thái đang dừng và chờ worker kết thúc. Không tự kết thúc cưỡng ép tiến trình sao chép.
- Nếu bị tắt máy hoặc khởi động lại, báo cáo/batch cuối chưa flush có thể thiếu một số kết quả. Đối chiếu PST trước khi xuất lại. Chưa có cơ chế tiếp tục xuyên phiên tự động hay bảo đảm không trùng bản sao xuyên phiên.
- PST trên ổ mạng và các thư mục OneDrive nhận diện qua biến môi trường bị từ chối. Chọn ổ đĩa cục bộ; không dùng thư mục đồng bộ khác.
- Không hỗ trợ New Outlook, Graph, POP/IMAP hoặc Calendar/Contacts/meeting request. Quét IPM.Note và lớp con trong mailbox Exchange.

## Các thay đổi về hiệu năng

- Một tiến trình STA và một namespace Outlook cho toàn bộ mỗi tác vụ; không mở tiến trình cho từng thư.
- PST chỉ được tìm/gắn một lần cho mỗi lần xuất.
- Bộ nhớ đệm thư mục đích theo các đoạn tên; không tìm lại cây thư mục cho từng thư.
- Quét bảng Outlook bằng GetTable/GetArray, 2.000 dòng mỗi lô. Không mở từng thư để tìm SMTP.
- Chỉ mở từng thư khi thực sự xuất. Không đọc lại cây nguồn hoặc lặp StoreID cho từng thư.
- Ghi kết quả theo 100 thư hoặc khoảng 500 ms; giao diện lấy tiến trình mỗi giây. CSV được ghi qua bộ đệm.
- Nhận diện thư bằng StoreID và EntryID; khi hai EntryID khác chuỗi, dùng Outlook CompareEntryIDs để xác nhận chúng là cùng thư. Không mở từng thư trong lúc quét chỉ để chuẩn hóa ID.
- Đường dẫn PST không chứa lớp thư mục tên tài khoản: Exchange Inbox/SPL → PST Inbox/SPL. Tên thư mục mặc định Exchange được chuẩn hóa theo ID của Outlook, gồm Inbox, Sent Items, Deleted Items, Drafts, Junk Email và Outbox. Ưu tiên thư mục Inbox đã có trong PST.
- COM không chạy song song để tránh truy cập Outlook từ nhiều luồng.

## Giao diện bubble 0.4

Gồm năm phần theo mẫu: quét email; thư mục dạng nút kèm số lượng; xem/chọn email với phân trang 100 dòng; ba lựa chọn xóa/PST có sẵn/PST mới; nhật ký. Có tìm kiếm, chọn/bỏ chọn trang hoặc toàn bộ kết quả đang lọc và tổng dung lượng đã chọn.

Nút Tải CSV ở phần quét xuất metadata của danh sách đã quét. Báo cáo thao tác là CSV riêng, có kết quả từng thư và các giai đoạn xóa. Nút Chọn file mở hộp thoại Windows khi lưu PST; vẫn có thể nhập đường dẫn.

### Lưu vào PST và tự xóa nguồn

1. Quét lại bằng bản 0.5.0, chọn email và PST có sẵn hoặc PST mới.
2. Chọn file PST, ví dụ 2026.pst. Exchange Inbox/SPL vào PST Inbox/SPL, không qua thư mục tài khoản.
3. Giữ tùy chọn tự xóa bật, bấm thực hiện và xác nhận một lần trước khi bắt đầu. Không hỏi lại sau khi sao chép.
4. Ứng dụng sao chép và xác minh PST trước. Sau đó gọi DeleteMessages với DELETE_HARD_DELETE cho đúng các ID nguồn có biên nhận, rồi kiểm tra chúng đã vắng khỏi thư mục nguồn. Lệnh này bỏ qua bước chuyển vào Deleted Items để giảm thời gian.
5. Đối chiếu số bản PST đã kiểm tra, số ID đã xóa và báo cáo CSV. Thư sao chép lỗi hoặc thiếu biên nhận sẽ không bị xóa; lỗi/partial delete không tự retry.

Để kiểm tra thủ công trước khi xóa, bỏ chọn tự xóa trước khi chạy. Nếu dừng trong giai đoạn sao chép, không bắt đầu tự xóa. Trong giai đoạn xóa, nút dừng chặn các bước tiếp theo; những thư đã xóa không thể hoàn tác qua công cụ. Không tự tiếp tục tác vụ sau khi mở lại ứng dụng. Biên nhận và báo cáo vẫn nằm trong thư mục tác vụ. Các thư mục sai tên tài khoản do bản cũ tạo không bị tự xóa.
Dữ liệu và nhật ký mặc định lưu trong **%AppData%\KTG\OutlookMailManager**. Nhật ký tổng là operations.log; mỗi tác vụ có thư mục riêng với request, trạng thái, batch và report.csv. Giữ các tệp này riêng tư vì chứa định danh/metadata email. Khi chạy kiểm thử, dùng --data-dir để tách dữ liệu giả khỏi dữ liệu người dùng.

Đã kiểm thử thật luồng quét và sao chép 5 email vào PST thử riêng: 5 bản nằm đúng Inbox/SPL, mở lại được, thư nguồn vẫn đủ 5. PST 2026 của người dùng không bị ghi trong phép thử. Bản 0.4.3 đã xử lý đúng 5 thư nguồn còn trong Deleted Items của lần xuất người dùng: đối chiếu bản PST, định danh và nội dung văn bản/HTML trước khi xóa. Đọc lại bảng Deleted Items xác nhận còn 0 thư. Không dọn thư khác và không xóa PST.

Vẫn cần thao tác COM cho từng bản sao. Chưa triển khai Extended MAPI bulk copy: cần kiểm thử provider, bitness Outlook và kết quả từng thư khi có partial completion. Không có nút bật một đường bulk chưa được kiểm chứng.

## Build EXE Rust

Đã cài Rust/Cargo 1.99.1, rustfmt, Visual Studio Build Tools 2022 17.14.45 và Windows SDK 10.0.26100 trên máy này. Bản 0.4 có 6 Rust unit tests, bao gồm đóng ứng dụng khi worker đang chạy. Xem TEST-RESULTS.txt để biết các kiểm thử đã chạy và các phần chưa đo với Outlook thật.

Bộ cài Microsoft trả mã 3010 (khuyến nghị khởi động lại Windows). Việc biên dịch và chạy kiểm thử EXE đã thành công trong phiên hiện tại; không tự khởi động lại máy.

Để build:

1. Cài Rust stable từ [trang chính thức](https://rust-lang.org/tools/install/).
2. Với toolchain Windows MSVC, cài Microsoft C++ Build Tools và Windows SDK theo [hướng dẫn Microsoft](https://learn.microsoft.com/en-us/windows/dev-environment/rust/setup).
3. Mở PowerShell trong thư mục dự án, chạy:

    powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1

Build chạy Rust unit tests, bộ kiểm thử PowerShell, tạo KTG-Outlook-Mail-Manager-Desktop-0.5.23.exe trong thư mục gốc rồi kiểm tra máy chủ trên EXE vừa build. Cửa sổ dùng tao 0.37.2 và wry 0.57.2, phiên bản phụ thuộc được cố định trong Cargo.lock. Dùng Rust stable; đã kiểm tra với Rust 1.99.1.

Có thể chạy riêng:

    cargo test --locked
    cargo build --release --locked
    powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\test-core.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\test-scan.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\test-server.ps1

tests/test-ui.cjs là kiểm thử trình duyệt với dữ liệu giả, yêu cầu Node và Playwright. Biến KTG_CHROME chỉ định đường dẫn Chrome; NODE_PATH có thể trỏ tới thư mục module được cài. Không cần những công cụ này để dùng Start.cmd.

Để kiểm tra EXE, dùng tham số -ServerExecutable với tests/test-server.ps1. Với test-ui.cjs, đặt biến KTG_SERVER_EXE trỏ tới EXE để kiểm thử các tài nguyên nhúng trong bản Rust thay cho máy chủ PowerShell.

## Cấu trúc nguồn

- src/main.rs: HTTP server Rust chỉ lắng nghe 127.0.0.1, token ngẫu nhiên, một worker, hủy mềm.
- src/desktop.rs: cửa sổ Windows, WebView2 riêng, chặn điều hướng ra ngoài và đợi worker khi đóng.
- scripts/server.ps1: máy chủ thay thế để chạy trên Windows chưa có Rust.
- scripts/bridge.ps1: kết nối Outlook, quét cache và mở PST.
- scripts/core.ps1: sao chép/kiểm tra, cache thư mục, tiến trình và báo cáo.
- web/: giao diện tiếng Việt, tìm kiếm, phân trang, chọn thư và CSV.
- tests/: kiểm thử mô phỏng, lỗi, quét và máy chủ.
- data/session-*/job-*/: dữ liệu tác vụ được tạo khi chạy. Mỗi tác vụ có request.json, state.json, các batch và report.csv nếu xuất.

Các file trong data có thể chứa metadata thư. Chỉ lưu cục bộ, không đưa vào Git hoặc gửi kèm khi chia sẻ nguồn. Cần lưu riêng báo cáo trước khi chuyển sang tác vụ khác; các file cũ vẫn còn trên đĩa.

## Benchmark trên hộp thư thật

Đã quét thật 1.358 email cache và thử sao chép 5 email vào PST thử riêng. Chưa benchmark 40.000 email. Tệp TEST-RESULTS.txt phân biệt kết quả thật và mô phỏng.

Khi benchmark trên dữ liệu thật: bỏ chọn tự xóa, thử 100 thư rồi 1.000 thư trước; ghi số thư, tổng dung lượng, thời gian, lỗi và report.csv. Kiểm tra bản PST mở được, đọc nội dung/tệp đính kèm và xác nhận thư nguồn còn nguyên. Sau đó mới thử 40.000 thư. Đo trên cùng mailbox/cache và ổ đĩa, dùng PST mới khác nhau cho mỗi lần. Không dùng chế độ Move của EXE cũ để lấy baseline vì nó tác động thư gốc.

Không suy ra thời gian xuất PST từ benchmark mô phỏng; không cam kết mục tiêu 15–30 phút khi chưa đo thật.

## Tài liệu API tham chiếu

- [MailItem.Copy](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.copy): ví dụ sao chép rồi chuyển bản sao sang thư mục khác.
- [DownloadState](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.oldownloadstate?view=outlook-pia): giá trị full item là 1.
- [IMAPIFolder::CopyMessages](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imapifolder-copymessages): bulk copy có kết quả partial completion; MESSAGE_MOVE chuyển thư nếu được đặt. API này chưa được dùng trong bản 0.3.
- [CompareEntryIDs](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._namespace.compareentryids?view=outlook-pia): hai chuỗi EntryID khác nhau có thể đại diện cho cùng một thư; cần so sánh qua Outlook.

### Sửa PST mới trong bản 0.5.0

Không dùng Store.GetDefaultFolder để chọn Inbox của PST mới. Trên máy này API trả về một Inbox nằm ngoài cây thư mục hiển thị: email vẫn mở được bằng EntryID nhưng không thấy trong Outlook. Mọi đường dẫn đích nay được tìm/tạo trực tiếp từ GetRootFolder().Folders. Trước khi tự xóa nguồn, ứng dụng đi lại đường dẫn từ gốc PST và đối chiếu thư mục chứa bản sao; không chỉ kiểm tra EntryID của bản sao.

Đã đưa 5 thư của lần xuất PST 2027 về đúng 2027/Inbox/SPL, có lưu .msg dự phòng riêng trước khi sửa và không xóa thêm thư. Đã kiểm thử tạo PST thử hoàn toàn mới với 5 thư thật, không tạo Inbox trước: 5 bản nằm trong cây hiển thị và mở lại được, 5 thư nguồn giữ nguyên, tắt tự xóa trong phép thử. Không đưa email/bản dự phòng riêng tư vào gói phát hành.
## Tối ưu quét 0.4.6

- Đọc GetArray 2.000 dòng/lô; không mở từng MailItem để quét.
- Chuyển đổi và ghi JSON trực tiếp theo lô bằng bộ xử lý .NET nhúng, tránh vòng lặp PowerShell và ConvertTo-Json trên từng đối tượng thư. Kiểm tra dừng mỗi 100 dòng; giữ số thư đã ghi khi dừng.
- Bỏ thiết lập cột bảng trên thư mục trống. Tính trước tên/đường dẫn thư mục dùng chung cho lô.
- Giao diện tải tối đa 4 batch song song, đếm thư mục tăng dần, tái sử dụng kết quả lọc và chỉ dựng 100 hàng/trang.

Đo trên máy này: 70.000 dòng mô phỏng xử lý/ghi metadata giảm 9,119 → 3,771 giây (khoảng 2,4 lần). Đọc bảng giảm 70 → 35 lần. Không phải tốc độ Outlook thật. Giao diện tải 70.000 thư/100 thư mục mô phỏng trong 2,161 giây, tìm kiếm/lọc/chọn tất cả đạt.

Quét chỉ đọc trên Outlook thật: 1.352 thư, worker 3,02 → 2,50 giây; tổng khởi động+worker 3,598 → 3,123 giây. Cùng định danh, số thư và metadata; không sửa/xóa email. Đây là một cặp đo trước/sau, không phải cam kết cho máy khác. Chưa có hộp thư thật 70.000 email để đo; tốc độ còn phụ thuộc OST, số thư mục, ổ đĩa và Outlook.

Có thể chạy tests/test-scan-benchmark.ps1 -Count 70000 để đo mô phỏng. tests/test-large-ui.cjs kiểm tra 70.000 thư giả, yêu cầu KTG_SERVER_EXE, Node/Playwright.
## Tối ưu xuất PST và xóa nguồn 0.4.7

Giữ nguyên ánh xạ Inbox/SPL → PST Inbox/SPL, cả PST mới/có sẵn. Một namespace và một PST được tái sử dụng trong toàn tác vụ.

- Giữ biên nhận do chính worker tạo trong bộ nhớ khi tự xóa, tránh đọc/parse lại toàn bộ JSON của 70.000 bản sao. Biên nhận đầy đủ vẫn ghi ra đĩa trước khi bắt đầu xóa. Luồng xóa sau xuất thủ công vẫn đọc biên nhận từ đĩa.
- Tái sử dụng kết quả kiểm tra cây thư mục trong tối đa 500 ms hoặc 100 lần dùng, rồi kiểm tra lại. Cache chỉ thuộc lần cleanup hiện tại. Mỗi email vẫn được mở lại riêng trong PST, kiểm tra ID/store/folder/subject/attachments/Internet ID trước khi xóa nguồn. Không thao tác di chuyển/đổi tên thư mục PST trong khi tác vụ đang chạy.
- Đọc Root.EntryID một lần, tái sử dụng EntryID nguồn đã đọc; dùng khóa cache giữ nguyên ranh giới tên thư mục.
- Ghi CSV bằng bộ xử lý .NET theo bộ đệm. Vẫn flush audit trước khi chuyển nguồn và trước khi Delete, gồm ID mới trong Deleted Items. Kết quả cuối flush theo lô/thời gian; không tự thử lại các mục lỗi.
- Giao diện không giữ thêm bản sao biên nhận kiểm tra trong cache kết quả, giảm bộ nhớ cho danh sách lớn.

Một cặp benchmark tuần tự 2.000 thư giả qua đủ copy/kiểm tra/xóa: 51,811 → 39,441 giây (~24% giảm thời gian). Đây là mô phỏng, không phải tốc độ Outlook thật. Lần đo đếm có gắn instrumentation: lượt đọc cây đích giảm 6.011 → 1.991; số lần mở từng email để kiểm tra vẫn giữ nguyên. Thời gian thực phụ thuộc COM, OST/PST, dung lượng thư và ổ đĩa.

Đã thử sao chép thật 5 thư vào PST mới riêng: 5 bản nằm trong đường dẫn hiển thị đúng, mở lại được, 5 nguồn còn nguyên. Phần tự xóa của lần cập nhật này chỉ thử bằng đối tượng giả. Chưa benchmark xuất/xóa 70.000 thư thật và chưa triển khai bulk MAPI. Không suy ra thời gian hoàn thành 70.000 thư từ mô phỏng.

Có thể chạy tests/test-export-benchmark.ps1 -Count 2000 hoặc tăng Count cho phép thử mô phỏng lớn hơn. Phép thử tạo đầy đủ CSV/biên nhận và đối chiếu các bản sao, không kết nối Outlook.
## MAPI theo lô 0.5.0

Chọn **Chuyển MAPI theo lô 100 thư** (bật sẵn) khi xuất PST. Hiện hỗ trợ Outlook Classic 64-bit. Bộ DLL native được nhúng trong EXE Rust; không cần cài Redemption hoặc add-in. Nếu máy dùng Outlook 32-bit hoặc provider không hỗ trợ, tắt tùy chọn này để dùng luồng cũ; không tự retry lô đã gọi CopyMessages.

Ứng dụng nhóm theo mailbox, folder nguồn và đường dẫn đích. Trước khi gọi native, từng nguồn phải có DownloadState đầy đủ cache. Mỗi lô giao diện 1.000 thư được gửi thành một lệnh MAPI trực tiếp từ Exchange vào đúng thư mục PST đã chọn; trước/sau khi sao chép đối chiếu EntryID, metadata và thư mục cha để chỉ nhận diện các mục mới của lô hiện tại. Không tạo thư mục KTG-batch và không dùng MESSAGE_MOVE với nguồn Exchange. Nếu copy dở hoặc xác minh không khớp, giữ nguồn; các bản sao chưa xác minh vẫn nằm trong thư mục đích và được ghi lý do trong manifest native-batch-*.json.

Sau đó luồng kiểm tra hiện có mở lại từng bản PST, đối chiếu nguồn/cache/timestamp/dung lượng/tệp đính kèm và đường dẫn hiển thị. Chỉ thư kiểm tra đạt mới được chuyển nguồn vào Deleted Items rồi mở lại, Delete và xác nhận không còn truy cập được. Hoàn thành từng lô rồi mới sang lô kế tiếp; khi dừng, các lô trước có thể đã xóa nguồn thành công, lô chưa kiểm tra giữ nguồn. Biên nhận và CSV nối tiếp xuyên lô, tổng tiến trình là toàn tác vụ.

**Tạm dừng / Tiếp tục** dừng ở ranh giới an toàn; lệnh MAPI đang chạy cần hoàn thành trước. Có thể hủy trong khi đang tạm dừng. Phiên MAPI riêng không gửi/nhận thư, không đổi Work Offline của Outlook. Quét cache và xuất MAPI được phép khi Outlook vẫn trực tuyến; từng thư vẫn phải có DownloadState đầy đủ để được kiểm tra/xóa. Chế độ xuất cũ và xóa riêng vẫn yêu cầu Work Offline. Trong lúc kiểm tra/xóa, công cụ vẫn gọi Outlook cho từng thư, nên không cam kết Outlook hoàn toàn không chậm; dùng Tạm dừng nếu cần.

Đã thử trên dữ liệu thật: 5 thư (có tệp đính kèm) trong một lần gọi MAPI copy khoảng 0,26–0,28 giây, đối chiếu Body/Subject/attachment count và không còn folder tạm khi thành công. Một lô 50 thư qua worker thật: 50 bản mở lại được, đúng đường dẫn và 50 nguồn giữ nguyên. Tất cả phép thử thật đều tắt tự xóa; automatic cleanup thử bằng đối tượng giả. Chưa benchmark chuyển/xóa 40.000 thư thật, không suy ra thời gian toàn tác vụ từ 0,26 giây copy của 5 thư.

Build cần chạy native/build.cmd (MSVC 64-bit) trước cargo. build.ps1 đã chạy bước này và test-bulk.ps1. Header MAPI chính thức Microsoft được vendored kèm giấy phép trong native/. Không có dữ liệu thư thật trong gói phát hành.
## Sửa lỗi phát hiện khi thử 1.000 thư — 0.5.1

Một thư không tiêu đề có thể không có PR_SUBJECT_W. Bản 0.5.0 coi đây là lỗi thuộc tính bắt buộc và bỏ qua cả lô 50. Bản 0.5.1 chỉ coi MAPI_E_NOT_FOUND của Subject là chuỗi rỗng; vẫn đối chiếu subject nguồn/đích, search key, ID, đường dẫn, Internet ID, số tệp, dung lượng và thời điểm sửa. Lỗi quyền hoặc lỗi đọc khác không bị bỏ qua. Không tự retry các lô đã thất bại.

Phép thử thật 1.000 thư/719 MiB trên cùng hộp thư và cùng ổ D: 0.4.7 sao chép đủ trong 142,23 giây; 0.5.0 chỉ đạt 950 do lỗi trên; 0.5.1 sao chép đủ 1.000 trong 111,17 giây, 21 lô. Đây là đo phần copy và kiểm tra metadata; kiểm tra sâu body/HTML/tệp đính kèm phục vụ thử nghiệm tính riêng. Không suy ra tốc độ 40.000 thư từ một lần đo, do bản cũ chạy trước và cache có thể ảnh hưởng.
Đợt thử 1.000 thư đã hoàn tất: hai PST đầy đủ qua kiểm tra sâu, sau đó xóa 1.000 nguồn và các ID tương ứng trong Deleted Items, 0 lỗi. Mở lại đủ 1.000 bản ở mỗi PST sau xóa. Cache 1.352 → 352; Outlook vẫn Offline, chưa xác minh đồng bộ Exchange.

## Dọn nguồn theo lô — 0.5.2

Sau khi từng lô tối đa 100 thư đã được sao chép và biên nhận PST được kiểm tra, MAPI đối chiếu lại nguồn, đường dẫn PST hiển thị và từng bản sao trước khi thay đổi nguồn. Thư nguồn được chuyển vào Deleted Items theo lô; ứng dụng ghi bền vững các ID mới quan sát được rồi mới gọi xóa danh sách ID đó trong Deleted Items. Không dùng EmptyFolder hoặc DELETE_HARD_DELETE. Lệnh xóa không tự chạy lại khi kết quả một phần hoặc không xác định; các ID còn hiện diện được báo lỗi và giữ nguyên. Việc giữ lại trong Recoverable Items vẫn tùy chính sách Exchange.

Phép thử tổng hợp trên PST riêng: 1.000 thư được dọn bằng 20 lô trong 3,51 giây; luồng Outlook từng thư mất 124,83 giây (xấp xỉ nhanh hơn 35,6 lần ở phép thử này). 1.000 bản sao PST vẫn còn nguyên và một thư ngoài lô trong Deleted Items được giữ lại. Đây không phải phép đo trên Exchange hoặc 40.000 thư; thời gian thật còn tùy Exchange, cache, dung lượng và chính sách giữ thư. Bản 0.5.2 giữ kiểm tra Outlook Work Offline cho thao tác đọc/xóa Exchange và không tự bật chế độ đó.
## Giảm lượt Outlook COM sau copy — 0.5.3

MAPI theo lô đã xác minh ID nguồn và bản sao trong đúng thư mục PST, nhưng 0.5.2 vẫn mở lại từng bản sao và lặp lại các kiểm tra Outlook COM sau đó. Bản 0.5.3 dùng biên nhận MAPI của lô để tạo kết quả và biên nhận PST trực tiếp, đồng thời gộp ghi trạng thái theo lô. Vẫn kiểm tra DownloadState trước khi copy; MAPI chỉ đọc nguồn Exchange ở chế độ cache-only; PST phải khớp theo search key, tiêu đề, Internet Message ID, số tệp và thư mục đích. Chỉ các biên nhận đó mới đủ điều kiện cho bước xóa nguồn.

Thay đổi này loại bỏ lượt mở lại từng nguồn/bản sao sau copy và tránh ghi state.json sau mỗi thư. Phần còn lại của tốc độ copy phụ thuộc thời gian đọc cache Exchange, thao tác MAPI trong từng lô và kích thước thư. Chưa đo lại 1.000 thư Exchange thật sau thay đổi này; cần đối chiếu trên mailbox/cache thực trước khi ước tính 40.000 thư.
Benchmark sao chép mới (PST tổng hợp): 1.000 thư qua MAPI copy theo lô và xử lý biên nhận mất 12,31 giây, sau đó bước dọn nguồn mất 3,13 giây. Phép thử không dùng Exchange thật và không so trực tiếp với mốc copy 111,17 giây trước đó vì mailbox/provider khác nhau. Đây là xác nhận luồng MAPI và biên nhận chạy được, không phải dự báo thời gian 40.000 thư thật.
## Tối ưu quét cache — 0.5.4

Tăng số dòng Outlook lấy mỗi lượt GetTable/GetArray từ 2.000 lên 5.000. Với 70.000 thư giả lập, số lượt đọc bảng giảm từ 35 xuống 14; thời gian encoder giảm từ 2,198 xuống 2,053 giây (khoảng 6,6%). Đây là benchmark bảng/encoder giả lập, không mô phỏng độ trễ Exchange hoặc OST thật. Kích thước lô tối đa vẫn giới hạn ở 5.000 để giữ các tệp kết quả và cập nhật giao diện ở mức vừa phải.

Bản 0.5.4 không đổi cách lọc cache Exchange, bộ lọc ngày, phân loại thư, người gửi hoặc định dạng kết quả. Chưa có phép đo mailbox thật sau tối ưu; mốc tham khảo trước đó là 1.300 thư trong 4,7 giây, còn 1.000 thư khoảng 3,6 giây nếu quy đổi tuyến tính.
## Quét theo lô 1.000 dòng — 0.5.5

Outlook GetTable/GetArray được đọc tối đa 1.000 dòng mỗi lượt. Encoder ghi từng lô ra danh sách kết quả ngay khi đọc xong; có kiểm tra dừng trong mỗi lô và định kỳ mỗi 100 dòng. Điều này giới hạn kích thước lô và giúp cập nhật danh sách từng phần. Các bộ lọc Exchange cache-only, ngày, thư IPM.Note và metadata giữ nguyên.

Lô nhỏ hơn làm tăng số lần gọi GetArray so với 0.5.4 (5.000 dòng/lượt), nên tổng thời gian có thể tăng đôi chút; lợi ích là mỗi lượt ngắn hơn và kết quả/dừng tác vụ có hạt mịn hơn. Chưa đo lại tốc độ trên mailbox thật sau thay đổi.

## Sao chép theo lô 1.000 thư — 0.5.9

Một lô ngoài gồm tối đa 1.000 thư cùng thư mục đích. Kiểm tra cache và MAPI CopyMessages chạy thành các nhóm 100 do giới hạn provider/API; giao diện cập nhật tiến độ trong khi xử lý. Sau khi biên nhận và toàn bộ bản sao đủ điều kiện của lô được kiểm tra, dọn nguồn được chia thành nhóm tối đa 100, với kiểm tra biên nhận riêng cho từng thư. Hủy trước khi kiểm tra xong sẽ giữ nguồn; các bản sao không được xác minh không bị xóa khỏi Exchange. 

Lô 1.000 giảm số lần xử lý vòng ngoài; không đồng nghĩa một lệnh COM/MAPI khổng lồ. PST thử nghiệm tổng hợp 1.000 thư: copy MAPI 6,96 giây, cleanup 2,52 giây; 1.000 bản còn đủ trong PST và sentinel ngoài lô được giữ. Đây không phải số đo mailbox Exchange thật.

## Giảm lượt kiểm tra lặp khi xuất PST — 0.5.9

Đã bỏ lượt kiểm tra DownloadState/GetItemFromID riêng qua Outlook COM trước mỗi nhóm MAPI. Native MAPI vốn đã mở từng nguồn bằng cờ cache-only và đối chiếu loại thư, thư mục, metadata; giờ kiểm tra đó cũng quyết định thư nào đủ điều kiện sao chép. Thư không mở được từ cache được ghi lỗi riêng, các thư khác trong nhóm vẫn tiếp tục. Không retry danh sách sau khi CopyMessages đã bắt đầu. 

Ở bước cleanup, PST đã được mở và xác minh ngay trước khi chuyển thư nguồn vào Deleted Items; bước xóa tiếp tục xác minh lại chính xác thư trong Deleted Items và chỉ xóa đúng các EntryID đã quan sát. Bỏ lần mở PST lặp thứ ba cho từng email. 

Mục tiêu 40.000 thư trong 10 phút chưa được xác nhận. Cần đo mailbox Exchange thật: độ trễ OST/MAPI và kích thước thư có thể chi phối; PST tổng hợp không dự báo được thời gian đó.

## Sửa lỗi lô MAPI bị hủy bởi ID xung đột — 0.5.9

Kết quả 0x80070057 trước đây được gộp chung, dù native worker đã biết lý do chi tiết. Bản này ghi nguyên nhân HRESULT vào báo cáo và loại riêng các thư có Search Key trùng khỏi nhóm, giữ lại mọi bản nguồn xung đột, đồng thời tiếp tục sao chép những thư có định danh rõ ràng trong cùng nhóm. Lỗi mở/đối chiếu từng thư cũng được ghi theo ID, thay vì bỏ cả lô. Chỉ bản PST có biên nhận hợp lệ mới vào danh sách dọn nguồn.

## Xử lý Search Key trùng khi dọn nguồn — 0.5.9

Các biên nhận PST hợp lệ có Search Key trùng không còn bị từ chối cả nhóm. Các khóa duy nhất vẫn được chuyển vào Deleted Items theo lô; khóa bị trùng được xử lý từng EntryID một, rồi xác minh chính xác một mục mới xuất hiện và đúng metadata trước khi xóa vĩnh viễn. Mục nào mơ hồ hoặc sai metadata vẫn được giữ nguồn. 

Luồng copy cũng ghi lý do từng thư không xuất hiện trong thư mục PST đích, không qua xác minh metadata, hoặc không tìm thấy đúng một bản ở thư mục cuối. Việc này giúp nhận biết vì sao thư không được phép dọn nguồn.

## Giữ lý do từng thư khi gộp kết quả MAPI — 0.5.10

Các kết quả từng nhóm CopyMessages giờ giữ lại cả danh sách thư bị bỏ qua và lý do xác minh thất bại, cùng HRESULT/lỗi của nhóm. Bảng CSV sẽ nêu lý do theo từng ID; lỗi một thư không dừng xác minh thư còn lại. Mục không có biên nhận PST vẫn giữ nguyên ở Exchange. Bản này cải thiện chẩn đoán 931 thư không được xác minh trong lần chạy 0.5.9; cần chạy lại nhóm thử nhỏ để xác nhận lỗi provider cụ thể.

## Ghép bản sao khi Search Key trùng — 0.5.15

Không còn bỏ cả thư có Search Key trùng. Trong mỗi lô MAPI, ứng dụng gom nguồn theo Search Key và chữ ký metadata (tiêu đề, Internet Message ID, kích thước, số tệp đính kèm). Trước khi cấp biên nhận PST, số bản mới tại thư mục đích của từng nhóm phải khớp chính xác với số thư nguồn; kiểm tra riêng từng bản và recheck từng nguồn. Chỉ nhóm đủ số lượng và khớp metadata mới đủ điều kiện dọn. Nhóm thiếu/thừa/không đọc được sẽ giữ nguồn và ghi nguyên nhân.
## Đối chiếu PST không phụ thuộc Search Key — 0.5.15

Exchange có thể đổi PR_SEARCH_KEY khi chép sang PST. Phiên bản này chụp EntryID thư mục đích trước khi chép, duyệt toàn bộ thư mục PST đích để phát hiện các bản mới, rồi so nhóm theo tiêu đề, Internet Message ID, kích thước và số tệp đính kèm. Chỉ nhóm có đủ số lượng mới ở thư mục PST đích mới nhận biên nhận cho phép dọn nguồn. Nếu metadata không đủ hoặc số lượng khác nhau, giữ thư nguồn. Việc duyệt thư mục PST đầy đủ tăng lượng đọc MAPI, nên cần benchmark trên Outlook thực tế trước khi xử lý hàng chục nghìn thư.

## Đọc EntryID PST kể cả khi Search Key không được cung cấp — 0.5.15

Bảng nội dung PST có thể không trả PR_SEARCH_KEY. Khi liệt kê toàn thư mục, phiên bản này chỉ yêu cầu PR_ENTRYID; việc xác minh bản sao PST cho phép thiếu Search Key, nhưng vẫn yêu cầu đọc được metadata khác, đúng thư mục và đủ số lượng. Việc đối chiếu nguồn Exchange vẫn yêu cầu Search Key và kiểm tra nguồn không đổi.

## Ghép bản sao PST bằng Internet Message ID — 0.5.15

PST có thể đổi Search Key và kích thước provider báo sau khi chép. Khi thư có Internet Message ID, đối chiếu theo ID + tiêu đề + số tệp đính kèm; khi không có ID thì vẫn yêu cầu Search Key + tiêu đề + kích thước + số tệp đính kèm khớp. Dọn nguồn cũng xác minh Internet Message ID với PST; thư không có ID giữ tiêu chí Search Key nghiêm ngặt. Báo lỗi giờ ghi số thư nguồn, số khớp và tổng thư mới đọc được ở thư mục PST đích để xác định bước sai. Nếu không ghép chắc chắn, nguồn được giữ.


## Ghi thư trực tiếp vào thư mục PST đã chọn — 0.5.20

Không còn tạo KTG-batch-{GUID} hay chạy thêm bước chuyển thư từ folder tạm sang Inbox\abc. MAPI copy thẳng vào folder đích, so sánh EntryID trước/sau, xác minh số bản mới theo chữ ký metadata và xác minh thư mục cha. Chỉ bản sao xác minh đủ mới cấp biên nhận cho dọn nguồn. Lỗi hoặc kết quả không rõ thì giữ thư Exchange; bản sao chưa xác minh không bị xóa và có thể cần đối chiếu thủ công trong folder đích.



## Giảm quét lặp khi chuyển lô lớn — 0.5.21

Một lô 1.000 thư trước đây bị chia thành 10 lệnh native 100 thư; mỗi lệnh duyệt toàn thư mục PST trước và sau khi copy. Bản này gửi tối đa 1.000 thư trong một lệnh, bỏ một lượt duyệt PST cuối bị lặp và tái sử dụng ID đã xác minh để phát biên nhận. Vẫn đối chiếu toàn bộ mục mới, kiểm tra từng thư nguồn không đổi; nếu số lượng hoặc metadata không khớp, giữ nguồn. Đây là giảm số lần quét thư mục; chưa có benchmark Outlook thật cho bản này nên không cam kết thời gian mới trước khi đo thực tế.



## Giảm thao tác mở lại thư nguồn — 0.5.22

Worker MAPI vẫn chụp metadata nguồn trước khi copy và mở/đối chiếu từng bản mới trong PST. Worker không mở lại toàn bộ nguồn thêm lần nữa ngay sau copy; nếu bật tự dọn, giai đoạn dọn nguồn vẫn mở lại và so khớp thư nguồn với biên nhận PST trước mọi thao tác xóa. Nguồn thay đổi hoặc bản PST không khớp sẽ được giữ lại. Thay đổi này loại bỏ một lượt MAPI mở và đọc attachment trên mỗi thư trong giai đoạn copy. Chưa đo hiệu năng trên mailbox thật với phiên bản này.



## Tránh chờ lâu khi xử lý lô nhiều thư — 0.5.23

Ánh xạ biên nhận MAPI sang thư được chọn bằng dictionary theo EntryID, thay cho lọc lặp cả danh sách cho từng thư (O(n) thay cho O(n²)). Nhật ký giờ phân biệt đang chờ MAPI trả kết quả với bước xử lý biên nhận, thay vì ghi 100% trước khi lời gọi đồng bộ kết thúc. Không thay đổi điều kiện xác minh PST hoặc dọn nguồn.

