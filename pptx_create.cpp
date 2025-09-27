// C++ (Windows, MS Office COM) - Tạo PowerPoint giống nội dung đã có
// Yêu cầu: Windows + Microsoft PowerPoint cài sẵn (Office 2016+). 
// Bật / điều chỉnh đường dẫn thư viện nếu khác phiên bản Office.
// Build (MSVC): cl /EHsc /W4 pptx_create.cpp
// Chạy với quyền user bình thường.

// ----- IMPORT COM TYPELIBS (điều chỉnh phiên bản Office nếu cần) -----
// Với Office 2016 (Office16). Nếu bạn dùng Office15/Office14, đổi đường dẫn tương ứng.
#import "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\mso.dll" \
    rename_namespace("Office"), rename("RGB","MSO_RGB")
#import "C:\\Program Files\\Microsoft Office\\Root\\Office16\\MSPPT.OLB" \
    rename_namespace("PowerPoint")

#include <comdef.h>
#include <iostream>

int wmain() {
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        std::wcerr << L"CoInitialize failed: " << std::hex << hr << std::endl;
        return 1;
    }

    try {
        // Khởi tạo PowerPoint
        PowerPoint::_ApplicationPtr app;
        hr = app.CreateInstance(L"PowerPoint.Application");
        if (FAILED(hr)) {
            std::wcerr << L"Không khởi tạo được PowerPoint.Application (COM)." << std::endl;
            CoUninitialize();
            return 1;
        }
        app->Visible = Office::MsoTriState::msoTrue; // Cho phép xem khi tạo

        // Tạo Presentation mới
        PowerPoint::PresentationsPtr presSet = app->Presentations;
        PowerPoint::_PresentationPtr pres = presSet->Add(Office::MsoTriState::msoTrue);

        // Tiện ích: hàm thêm slide Title + Content
        auto addTitleAndContent = [&](long index, const wchar_t* title, const wchar_t* content) {
            PowerPoint::SlidesPtr slides = pres->Slides;
            PowerPoint::_SlidePtr s = slides->Add(index, PowerPoint::PpSlideLayout::ppLayoutText);
            // Placeholder 1: Title
            s->Shapes->Item(L"Title 1")->TextFrame->TextRange->Text = title;
            // Placeholder 2: Body
            s->Shapes->Item(L"Text Placeholder 2")->TextFrame->TextRange->Text = content;
            return s;
        };

        // ----- Slide 1: Title -----
        {
            PowerPoint::SlidesPtr slides = pres->Slides;
            PowerPoint::_SlidePtr s = slides->Add(1, PowerPoint::PpSlideLayout::ppLayoutTitle);
            s->Shapes->Item(1)->TextFrame->TextRange->Text = L"CƠ CHẾ GÂY BỆNH VÀ TÁC ĐỘNG LÊN VẬT CHỦ";
            s->Shapes->Item(2)->TextFrame->TextRange->Text = L"So sánh giữa Vi khuẩn và Virus";
        }

        // ----- Slide 2: Outline -----
        addTitleAndContent(2,
            L"Đề cương (Outline)",
            L"1. Giới thiệu chung về Vi khuẩn và Virus\n"
            L"2. Cơ chế gây bệnh của Vi khuẩn\n"
            L"3. Cơ chế gây bệnh của Virus\n"
            L"4. Tác động lên vật chủ (giáp mô, phổ xâm nhiễm, provirus)\n"
            L"5. Bảng so sánh tổng hợp\n"
            L"6. Phòng và trị bệnh\n"
            L"7. Kết luận"
        );

        // ----- Slide 3: Cơ chế gây bệnh Vi khuẩn -----
        addTitleAndContent(3,
            L"Cơ chế gây bệnh của Vi khuẩn",
            L"- Sản sinh ngoại độc tố (exotoxin) và nội độc tố (endotoxin - LPS).\n"
            L"- Capsule (giáp mô) chống thực bào, tăng độc lực.\n"
            L"- Enzyme phá hủy mô: hyaluronidase, collagenase.\n"
            L"- Có thể sống độc lập, tự biến dưỡng."
        );

        // ----- Slide 4: Cơ chế gây bệnh Virus -----
        addTitleAndContent(4,
            L"Cơ chế gây bệnh của Virus",
            L"- Chu trình 5 giai đoạn: Hấp phụ → Xâm nhập → Sao chép → Lắp ráp → Giải phóng.\n"
            L"- Phá hủy tế bào khi giải phóng (CPE).\n"
            L"- Tích hợp provirus vào NST (retrovirus: HIV).\n"
            L"- Không sản sinh độc tố hóa học."
        );

        // ----- Slide 5: Tác động lên vật chủ -----
        addTitleAndContent(5,
            L"Tác động lên vật chủ",
            L"- Vi khuẩn: gây viêm, sốt, nhiễm trùng huyết, tổn thương mô.\n"
            L"- Virus: chết tế bào, ung thư (HPV, HBV, HCV), nhiễm tiềm ẩn kéo dài.\n"
            L"- Giáp mô của vi khuẩn và envelope của virus giúp tránh hệ miễn dịch.\n"
            L"- Phổ xâm nhiễm: virus có thể hẹp (HIV) hoặc rộng (dại)."
        );

        // ----- Slide 6: Bảng so sánh -----
        addTitleAndContent(6,
            L"Bảng so sánh Vi khuẩn và Virus",
            L"| Tiêu chí | Vi khuẩn | Virus |\n"
            L"|---------|---------|-------|\n"
            L"| Cấu trúc | Tế bào hoàn chỉnh | Không có tế bào |\n"
            L"| Độc tố | Ngoại & nội độc tố | Không có, phá hủy tế bào |\n"
            L"| Sinh sản | Tự nhân lên | Phụ thuộc tế bào chủ |\n"
            L"| Phổ xâm nhiễm | Rộng | Hẹp hoặc rộng tùy loại |"
        );

        // ----- Slide 7: Phòng và trị bệnh -----
        addTitleAndContent(7,
            L"Phòng và trị bệnh",
            L"- Vi khuẩn: kháng sinh đặc hiệu, vaccine.\n"
            L"- Virus: vaccine phòng, thuốc kháng virus giới hạn (HIV, cúm).\n"
            L"- Vệ sinh môi trường, tăng đề kháng đặc hiệu và không đặc hiệu."
        );

        // ----- Slide 8: Kết luận -----
        addTitleAndContent(8,
            L"Kết luận",
            L"Vi khuẩn và Virus có cơ chế gây bệnh và tác động khác biệt.\n"
            L"Hiểu rõ cơ chế giúp phòng, chẩn đoán và điều trị hiệu quả."
        );

        // Lưu file (điều chỉnh đường dẫn theo máy của bạn)
        _bstr_t outPath(L"C:\\Temp\\Co_che_gay_benh_va_tac_dong_vat_chu.pptx");
        pres->SaveAs(outPath, PowerPoint::PpSaveAsFileType::ppSaveAsDefault);

        // Đóng và thoát
        pres->Close();
        app->Quit();
        CoUninitialize();

        std::wcout << L"Đã tạo file: " << (const wchar_t*)outPath << std::endl;
        return 0;
    }
    catch (const _com_error& e) {
        std::wcerr << L"COM error: " << e.ErrorMessage() << L" (HRESULT=" << std::hex << e.Error() << L")" << std::endl;
    }
    catch (const std::exception& ex) {
        std::wcerr << L"std::exception: " << _bstr_t(ex.what()) << std::endl;
    }

    CoUninitialize();
    return 1;
}
