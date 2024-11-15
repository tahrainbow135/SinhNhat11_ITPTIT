# Cài python và cài thư viện
# Lệnh cài thư viện: pip install pandas openpyxl
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

GMAIL_USER = "clb.it.ptit@gmail.com"
GMAIL_PASSWORD = "azlh vzma juaa brks"

# Đọc dữ liệu từ file Excel
# Đưa đường dẫn đến file Excel vào đây
df = pd.read_excel("E:/JetbrainsTool/PycharmProjects/CodePtit/SinhNhat11/Doanhnghiep.xlsx")

# Kết nối đến máy chủ Gmail
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(GMAIL_USER, GMAIL_PASSWORD)
successful_recipients = []

# Gửi email cho từng người
for index, row in df.iterrows():
    # Nhập email từ cột tên email
    email = row["Email"]
    # Nhập tên từ cột tên name
    name = row["Name"]

    # Lấy image_id từ cột chứa URL hoặc ID (tuỳ vào cách bạn lưu trong Excel)
    google_drive_link = row["ImageLink"]  # Giả sử cột chứa link là "ImageLink"

    # Trích xuất image_id từ URL Google Drive
    if "drive.google.com" in google_drive_link:
        if "/d/" in google_drive_link:
            image_id = google_drive_link.split("/d/")[1].split("/")[0]
        elif "id=" in google_drive_link:
            image_id = google_drive_link.split("id=")[1].split("&")[0]
        else:
            image_id = ""
    else:
        image_id = ""

    if image_id:
        image_url = f"https://drive.google.com/uc?export=view&id={image_id}"

        # Tạo email
        msg = MIMEMultipart("alternative")
        msg["From"] = GMAIL_USER
        msg["To"] = email
        msg["Subject"] = "[CLB IT PTIT] THƯ MỜI THAM DỰ SINH NHẬT 11 TUỔI"

        # Nội dung email thay name, time, location, link_form
        html_body = f"""
            <!DOCTYPE html>
            <html lang="en">
              <head>
                <meta charset="UTF-8" />
                <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                <title>Document</title>
              </head>
              <body>
                <div class="">
                  <div class="aHl"></div>
                  <div id=":15t" tabindex="-1"></div>
                  <div
                    id=":163"
                    class="ii gt"
                    jslog="20277; u014N:xr6bB; 1:WyIjdGhyZWFkLWY6MTgxNTcyOTg2ODg4MzkzNTk0MyJd; 4:WyIjbXNnLWY6MTgxNTcyOTg2ODg4MzkzNTk0MyIsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLG51bGwsbnVsbCxudWxsLDBd"
                  >
                    <div id=":164" class="a3s aiL">
                      <div dir="ltr">
                        <div>
                          <div class="adM"></div>
                          <table
                            id="m_9017395527194329753u_body"
                            style="
                              border-collapse: collapse;
                              table-layout: fixed;
                              border-spacing: 0;
                              vertical-align: top;
                              min-width: 320px;
                              margin: 0 auto;
                              background-color: #f9f9f9;
                              width: 100%;
                            "
                            cellpadding="0"
                            cellspacing="0"
                          >
                            <tbody>
                              <tr style="vertical-align: top">
                                <td
                                  style="
                                    word-break: break-word;
                                    border-collapse: collapse !important;
                                    vertical-align: top;
                                  "
                                >
                                  <div style="padding: 0px; background-color: transparent">
                                    <div
                                      style="
                                        margin: 0 auto;
                                        min-width: 320px;
                                        max-width: 600px;
                                        word-wrap: break-word;
                                        word-break: break-word;
                                        background-color: transparent;
                                      "
                                    >
                                      <div
                                        style="
                                          border-collapse: collapse;
                                          display: table;
                                          width: 100%;
                                          height: 100%;
                                          background-color: transparent;
                                        "
                                      >
                                        <div
                                          style="
                                            max-width: 320px;
                                            min-width: 600px;
                                            display: table-cell;
                                            vertical-align: top;
                                          "
                                        >
                                          <div
                                            style="
                                              background-color: #194781;
                                              height: 100%;
                                              width: 100% !important;
                                              border-radius: 0px;
                                            "
                                          >
                                            <div
                                              style="
                                                box-sizing: border-box;
                                                height: 100%;
                                                padding: 0px;
                                                border-top: 0px solid transparent;
                                                border-left: 0px solid transparent;
                                                border-right: 0px solid transparent;
                                                border-bottom: 0px solid transparent;
                                                border-radius: 0px;
                                              "
                                            >
                                              <table
                                                style="font-family: 'Cabin', sans-serif"
                                                role="presentation"
                                                cellpadding="0"
                                                cellspacing="0"
                                                width="100%"
                                                border="0"
                                              >
                                                <tbody>
                                                  <tr>
                                                    <td
                                                      style="
                                                        word-break: break-word;
                                                        padding: 10px;
                                                        font-family: 'Cabin', sans-serif;
                                                      "
                                                      align="left"
                                                    >
                                                      <table
                                                        width="100%"
                                                        cellpadding="0"
                                                        cellspacing="0"
                                                        border="0"
                                                      >
                                                        <tbody>
                                                          <tr>
                                                            <td
                                                              style="
                                                                padding-right: 0px;
                                                                padding-left: 0px;
                                                              "
                                                              align="center"
                                                            >
                                                              <img
                                                                align="center"
                                                                border="0"
                                                                src="https://ci3.googleusercontent.com/meips/ADKq_NabvMKZxBywT2clms1zM0WGv5dSBBVAAYov779nQUb2E7EEKAlP6HzVbhvH2HzxyVA00dbg0Mdw6btLeF-YFqYrYIX5=s0-d-e1-ft#https://share1.cloudhq-mkt3.net/d0c08cb94b4784"
                                                                alt=""
                                                                title=""
                                                                style="
                                                                  outline: none;
                                                                  text-decoration: none;
                                                                  clear: both;
                                                                  display: inline-block !important;
                                                                  border: none;
                                                                  height: auto;
                                                                  float: none;
                                                                  width: 100%;
                                                                  max-width: 580px;
                                                                "
                                                                width="580"
                                                                class="CToWUd a6T"
                                                                data-bit="iit"
                                                                tabindex="0"
                                                              />
                                                            </td>
                                                          </tr>
                                                        </tbody>
                                                      </table>
                                                    </td>
                                                  </tr>
                                                </tbody>
                                              </table>
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
            
                                  <div style="padding: 0px; background-color: transparent">
                                    <div
                                      style="
                                        margin: 0 auto;
                                        min-width: 320px;
                                        max-width: 600px;
                                        word-wrap: break-word;
                                        word-break: break-word;
                                        background-color: #ffffff;
                                      "
                                    >
                                      <div
                                        style="
                                          border-collapse: collapse;
                                          display: table;
                                          width: 100%;
                                          height: 100%;
                                          background-color: transparent;
                                        "
                                      >
                                        <div
                                          style="
                                            max-width: 320px;
                                            min-width: 600px;
                                            display: table-cell;
                                            vertical-align: top;
                                          "
                                        >
                                          <div
                                            style="
                                              background-color: #ffffff;
                                              height: 100%;
                                              width: 100% !important;
                                            "
                                          >
                                            <div
                                              style="
                                                box-sizing: border-box;
                                                height: 100%;
                                                padding: 0px;
                                                border-top: 4px solid #194781;
                                                border-left: 4px solid #194781;
                                                border-right: 4px solid #194781;
                                                border-bottom: 4px solid #194781;
                                              "
                                            >
                                              <table
                                                style="font-family: 'Cabin', sans-serif"
                                                role="presentation"
                                                cellpadding="0"
                                                cellspacing="0"
                                                width="100%"
                                                border="0"
                                              >
                                                <tbody>
                                                  <tr>
                                                    <td
                                                      style="
                                                        word-break: break-word;
                                                        padding: 20px 15px;
                                                        font-family: 'Cabin', sans-serif;
                                                      "
                                                      align="left"
                                                    >
                                                      <div
                                                        style="
                                                          font-size: 14px;
                                                          line-height: 160%;
                                                          text-align: center;
                                                          word-wrap: break-word;
                                                        "
                                                      >
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Kính gửi</span
                                                          ><strong
                                                            ><em
                                                              ><span
                                                                style="
                                                                  font-size: 14pt;
                                                                  font-family: 'Times New Roman',
                                                                    serif;
                                                                  color: #000000;
                                                                  line-height: 28.8px;
                                                                "
                                                              >
                                                              </span></em></strong
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >{name}</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >,&nbsp;</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            padding-top: 12pt;
                                                            padding-bottom: 12pt;
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Lời đầu tiên, cho phép </span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #ff9900;
                                                                line-height: 28.8px;
                                                              "
                                                              >CLB IT PTIT</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                          >
                                                            gửi tới </span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >{name}</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                          >
                                                            lời chào trân trọng và lòng biết
                                                            ơn sâu sắc vì sự đồng hành và
                                                            ủng hộ quý báu trong suốt thời
                                                            gian qua. Nhờ sự hỗ trợ nhiệt
                                                            tình từ Quý doanh nghiệp, </span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #ff9900;
                                                                line-height: 28.8px;
                                                              "
                                                              >IT PTIT</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                          >
                                                            đã có thể không ngừng phát
                                                            triển, trở thành điểm tựa và
                                                            nguồn cảm hứng cho nhiều thế hệ
                                                            sinh viên đam mê công
                                                            nghệ.</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Năm 2024 là một cột mốc đặc
                                                            biệt đối với chúng tôi. Nhìn lại
                                                            11 năm qua, CLB đã mang lại
                                                            nhiều giá trị về học tập lẫn
                                                            tinh thần tới các bạn sinh viên
                                                            trong và ngoài</span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #ff0000;
                                                                line-height: 28.8px;
                                                              "
                                                            >
                                                            </span></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >CLB.&nbsp;</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          &nbsp;
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Nhằm đánh dấu cột mốc quan
                                                            trọng này </span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #ff9900;
                                                                line-height: 28.8px;
                                                              "
                                                              >IT PTIT</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #ff9900;
                                                              line-height: 28.8px;
                                                            "
                                                          >
                                                          </span
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >rất vui mừng tổ chức </span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >lễ kỷ niệm 11 năm</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >. Sự kiện này là cơ hội để
                                                            chúng tôi nhìn lại những thành
                                                            tựu đã đạt được với lòng biết ơn
                                                            và hướng tới một tương lai đầy
                                                            hy vọng.</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          &nbsp;
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #ff9900;
                                                                line-height: 28.8px;
                                                              "
                                                              >IT PTIT</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #ff9900;
                                                              line-height: 28.8px;
                                                            "
                                                          >
                                                          </span
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >xin trân trọng kính mời </span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >{name}</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                          >
                                                            tới tham dự bữa tiệc Sinh nhật
                                                            với chủ đề </span
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >“Tinh Tú Vẫy Gọi”</span
                                                            ></strong
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >. </span
                                                          ><span
                                                            style="
                                                              font-size: 13pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 27.2px;
                                                            "
                                                            >Sự có mặt của Quý doanh nghiệp
                                                            là niềm vui và hạnh phúc của </span
                                                          ><span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >CLB</span
                                                          ><span
                                                            style="
                                                              font-size: 13pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 27.2px;
                                                            "
                                                            >.</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            text-align: justify;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            padding-top: 12pt;
                                                            padding-bottom: 12pt;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Thông tin chi tiết về buổi lễ
                                                            kỷ niệm:</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            text-align: justify;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            padding-bottom: 12pt;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >Thời gian: 17h30, Thứ Bảy,
                                                              ngày 23 tháng 11 năm
                                                              2024</span
                                                            ></strong
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            line-height: 160%;
                                                            text-align: justify;
                                                            background-color: rgb(
                                                              255,
                                                              255,
                                                              255
                                                            );
                                                            padding-bottom: 12pt;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >Địa điểm: Hội trường A2, Học
                                                              viện Công nghệ Bưu chính Viễn
                                                              thông (</span
                                                            ></strong
                                                          ><a
                                                            href="https://www.google.com/maps/dir/20.9752565,105.7852988/20.9807365,105.7874288/@20.9780266,105.7800735,16z/data=!3m1!4b1!4m4!4m3!1m1!4e1!1m0?entry=ttu&amp;g_ep=EgoyMDI0MTExMS4wIKXMDSoASAFQAw%3D%3D"
                                                            style="text-decoration: none"
                                                            target="_blank"
                                                            data-saferedirecturl="https://www.google.com/url?q=https://www.google.com/maps/dir/20.9752565,105.7852988/20.9807365,105.7874288/@20.9780266,105.7800735,16z/data%3D!3m1!4b1!4m4!4m3!1m1!4e1!1m0?entry%3Dttu%26g_ep%3DEgoyMDI0MTExMS4wIKXMDSoASAFQAw%253D%253D&amp;source=gmail&amp;ust=1731701344035000&amp;usg=AOvVaw1zof9U2MIzOUyx7gY3XjxQ"
                                                            ><strong
                                                              ><span
                                                                style="
                                                                  font-size: 14pt;
                                                                  font-family: 'Times New Roman',
                                                                    serif;
                                                                  color: #1155cc;
                                                                  text-decoration: underline;
                                                                  line-height: 28.8px;
                                                                "
                                                                >Km 10 Nguyễn Trãi, Hà Đông,
                                                                Hanoi, Vietnam</span
                                                              ></strong
                                                            ></a
                                                          ><strong
                                                            ><span
                                                              style="
                                                                font-size: 14pt;
                                                                font-family: 'Times New Roman',
                                                                  serif;
                                                                color: #000000;
                                                                line-height: 28.8px;
                                                              "
                                                              >)</span
                                                            ></strong
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            font-family: Cabin, sans-serif;
                                                            font-size: 14px;
                                                            white-space: normal;
                                                            color: #222222;
                                                            line-height: 160%;
                                                            background-color: #ffffff;
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Trân trọng,</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            font-family: Cabin, sans-serif;
                                                            font-size: 14px;
                                                            white-space: normal;
                                                            color: #222222;
                                                            line-height: 160%;
                                                            background-color: #ffffff;
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Lê Đức Hiếu</span
                                                          >
                                                        </p>
                                                        <p
                                                          style="
                                                            font-family: Cabin, sans-serif;
                                                            font-size: 14px;
                                                            white-space: normal;
                                                            color: #222222;
                                                            line-height: 160%;
                                                            background-color: #ffffff;
                                                            text-align: justify;
                                                            margin: 0px;
                                                          "
                                                        >
                                                          <span
                                                            style="
                                                              font-size: 14pt;
                                                              font-family: 'Times New Roman',
                                                                serif;
                                                              color: #000000;
                                                              line-height: 28.8px;
                                                            "
                                                            >Chủ nhiệm CLB IT PTIT</span
                                                          >
                                                        </p>
                                                      </div>
                                                    </td>
                                                  </tr>
                                                </tbody>
                                              </table>
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
            
                                  <div style="padding: 0px; background-color: transparent">
                                    <div
                                      style="
                                        margin: 0 auto;
                                        min-width: 320px;
                                        max-width: 600px;
                                        word-wrap: break-word;
                                        word-break: break-word;
                                        background-color: transparent;
                                      "
                                    >
                                      <div
                                        style="
                                          border-collapse: collapse;
                                          display: table;
                                          width: 100%;
                                          height: 100%;
                                          background-color: transparent;
                                        "
                                      >
                                        <div
                                          style="
                                            max-width: 320px;
                                            min-width: 600px;
                                            display: table-cell;
                                            vertical-align: top;
                                          "
                                        >
                                          <div
                                            style="
                                              background-color: #194781;
                                              height: 100%;
                                              width: 100% !important;
                                              border-radius: 0px;
                                            "
                                          >
                                            <div
                                              style="
                                                box-sizing: border-box;
                                                height: 100%;
                                                padding: 0px;
                                                border-top: 0px solid transparent;
                                                border-left: 0px solid transparent;
                                                border-right: 0px solid transparent;
                                                border-bottom: 0px solid transparent;
                                                border-radius: 0px;
                                              "
                                            >
                                              <table
                                                style="font-family: 'Cabin', sans-serif"
                                                role="presentation"
                                                cellpadding="0"
                                                cellspacing="0"
                                                width="100%"
                                                border="0"
                                              >
                                                <tbody>
                                                  <tr>
                                                    <td
                                                      style="
                                                        word-break: break-word;
                                                        padding: 10px;
                                                        font-family: 'Cabin', sans-serif;
                                                      "
                                                      align="left"
                                                    >
                                                      <table
                                                        width="100%"
                                                        cellpadding="0"
                                                        cellspacing="0"
                                                        border="0"
                                                      >
                                                        <tbody>
                                                          <tr>
                                                            <td
                                                              style="
                                                                padding-right: 0px;
                                                                padding-left: 0px;
                                                              "
                                                              align="center"
                                                            >
                                                              <img
                                                                align="center"
                                                                border="0"
                                                                src={image_url}
                                                                alt=""
                                                                title=""
                                                                style="
                                                                  outline: none;
                                                                  text-decoration: none;
                                                                  clear: both;
                                                                  display: inline-block !important;
                                                                  border: none;
                                                                  height: auto;
                                                                  float: none;
                                                                  width: 100%;
                                                                  max-width: 580px;
                                                                "
                                                                width="580"
                                                                class="CToWUd a6T"
                                                                data-bit="iit"
                                                                tabindex="0"
                                                              />
                                                            </td>
                                                          </tr>
                                                        </tbody>
                                                      </table>
                                                    </td>
                                                  </tr>
                                                </tbody>
                                              </table>
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
            
                                  <div style="padding: 0px; background-color: transparent">
                                    <div
                                      style="
                                        margin: 0 auto;
                                        min-width: 320px;
                                        max-width: 600px;
                                        word-wrap: break-word;
                                        word-break: break-word;
                                        background-color: #194781;
                                      "
                                    >
                                      <div
                                        style="
                                          border-collapse: collapse;
                                          display: table;
                                          width: 100%;
                                          height: 100%;
                                          background-color: transparent;
                                        "
                                      >
                                        <div
                                          style="
                                            max-width: 320px;
                                            min-width: 600px;
                                            display: table-cell;
                                            vertical-align: top;
                                          "
                                        >
                                          <div style="height: 100%; width: 100% !important">
                                            <div
                                              style="
                                                box-sizing: border-box;
                                                height: 100%;
                                                padding: 0px;
                                                border-top: 0px solid transparent;
                                                border-left: 0px solid transparent;
                                                border-right: 0px solid transparent;
                                                border-bottom: 0px solid transparent;
                                              "
                                            >
                                              <table
                                                style="font-family: 'Cabin', sans-serif"
                                                role="presentation"
                                                cellpadding="0"
                                                cellspacing="0"
                                                width="100%"
                                                border="0"
                                              >
                                                <tbody>
                                                  <tr>
                                                    <td
                                                      style="
                                                        word-break: break-word;
                                                        padding: 15px 20px 22px;
                                                        font-family: 'Cabin', sans-serif;
                                                      "
                                                      align="left"
                                                    >
                                                      <div align="center">
                                                        <div
                                                          style="
                                                            display: table;
                                                            max-width: 239px;
                                                          "
                                                        >
                                                          <table
                                                            align="center"
                                                            border="0"
                                                            cellspacing="0"
                                                            cellpadding="0"
                                                            width="25"
                                                            height="25"
                                                            style="
                                                              width: 25px !important;
                                                              height: 25px !important;
                                                              display: inline-block;
                                                              border-collapse: collapse;
                                                              table-layout: fixed;
                                                              border-spacing: 0;
                                                              vertical-align: top;
                                                              margin-right: 23px;
                                                            "
                                                          >
                                                            <tbody>
                                                              <tr
                                                                style="vertical-align: top"
                                                              >
                                                                <td
                                                                  align="center"
                                                                  valign="middle"
                                                                  style="
                                                                    word-break: break-word;
                                                                    border-collapse: collapse !important;
                                                                    vertical-align: top;
                                                                  "
                                                                >
                                                                  <a
                                                                    href="mailto:clb.it.ptit@gmail.com"
                                                                    title="Email"
                                                                    target="_blank"
                                                                  >
                                                                    <img
                                                                      src="https://ci3.googleusercontent.com/meips/ADKq_NbQ1awrA278bNOI6nn55qKwaX22pUwD-08j8vmXPUc8AUsalrrDsPGcCRFSNd4w7-QcV9i4F1_J1NgzWdaLkuYzaxxBBpNQd3uVfe52XVpI6js=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/email.png"
                                                                      alt="Email"
                                                                      title="Email"
                                                                      width="25"
                                                                      style="
                                                                        outline: none;
                                                                        text-decoration: none;
                                                                        clear: both;
                                                                        display: block !important;
                                                                        border: none;
                                                                        height: auto;
                                                                        float: none;
                                                                        max-width: 25px !important;
                                                                      "
                                                                      class="CToWUd"
                                                                      data-bit="iit"
                                                                    />
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
            
                                                          <table
                                                            align="center"
                                                            border="0"
                                                            cellspacing="0"
                                                            cellpadding="0"
                                                            width="25"
                                                            height="25"
                                                            style="
                                                              width: 25px !important;
                                                              height: 25px !important;
                                                              display: inline-block;
                                                              border-collapse: collapse;
                                                              table-layout: fixed;
                                                              border-spacing: 0;
                                                              vertical-align: top;
                                                              margin-right: 23px;
                                                            "
                                                          >
                                                            <tbody>
                                                              <tr
                                                                style="vertical-align: top"
                                                              >
                                                                <td
                                                                  align="center"
                                                                  valign="middle"
                                                                  style="
                                                                    word-break: break-word;
                                                                    border-collapse: collapse !important;
                                                                    vertical-align: top;
                                                                  "
                                                                >
                                                                  <a
                                                                    href="https://www.facebook.com/ITPTIT"
                                                                    title="Facebook"
                                                                    target="_blank"
                                                                    data-saferedirecturl="https://www.google.com/url?q=https://www.facebook.com/ITPTIT&amp;source=gmail&amp;ust=1731701344035000&amp;usg=AOvVaw0QVZoejOWdj4HKsEzs1ZDo"
                                                                  >
                                                                    <img
                                                                      src="https://ci3.googleusercontent.com/meips/ADKq_NbuJJEY9VDc4xerFh35zfwU6rXm9N4x-QL2sA79wKkpfySrsmgkmKJ7Afkx1b-PqBBbzaqf1i0g7ldsxkRq56yaANUi_JXNkBa7T7HNWfS-l-Uey5A=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/facebook.png"
                                                                      alt="Facebook"
                                                                      title="Facebook"
                                                                      width="25"
                                                                      style="
                                                                        outline: none;
                                                                        text-decoration: none;
                                                                        clear: both;
                                                                        display: block !important;
                                                                        border: none;
                                                                        height: auto;
                                                                        float: none;
                                                                        max-width: 25px !important;
                                                                      "
                                                                      class="CToWUd"
                                                                      data-bit="iit"
                                                                    />
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
            
                                                          <table
                                                            align="center"
                                                            border="0"
                                                            cellspacing="0"
                                                            cellpadding="0"
                                                            width="25"
                                                            height="25"
                                                            style="
                                                              width: 25px !important;
                                                              height: 25px !important;
                                                              display: inline-block;
                                                              border-collapse: collapse;
                                                              table-layout: fixed;
                                                              border-spacing: 0;
                                                              vertical-align: top;
                                                              margin-right: 23px;
                                                            "
                                                          >
                                                            <tbody>
                                                              <tr
                                                                style="vertical-align: top"
                                                              >
                                                                <td
                                                                  align="center"
                                                                  valign="middle"
                                                                  style="
                                                                    word-break: break-word;
                                                                    border-collapse: collapse !important;
                                                                    vertical-align: top;
                                                                  "
                                                                >
                                                                  <a
                                                                    href="https://www.youtube.com/channel/UC8Iwsz8PT07_yVpqEvG7MRw"
                                                                    title="YouTube"
                                                                    target="_blank"
                                                                    data-saferedirecturl="https://www.google.com/url?q=https://www.youtube.com/channel/UC8Iwsz8PT07_yVpqEvG7MRw&amp;source=gmail&amp;ust=1731701344035000&amp;usg=AOvVaw0GVV47BS77hYvJgQXcaEbu"
                                                                  >
                                                                    <img
                                                                      src="https://ci3.googleusercontent.com/meips/ADKq_NY06zLS_Qp0mU0LogYDcAPFvY3tHnEzKl1ZG2AzmOLekeQO-T6Qz-jdZKYYHyqVwflHbBZFHxNLIV8mQRErqSvYeTklqp7yTcaKa5N3AwZaQULSHg=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/youtube.png"
                                                                      alt="YouTube"
                                                                      title="YouTube"
                                                                      width="25"
                                                                      style="
                                                                        outline: none;
                                                                        text-decoration: none;
                                                                        clear: both;
                                                                        display: block !important;
                                                                        border: none;
                                                                        height: auto;
                                                                        float: none;
                                                                        max-width: 25px !important;
                                                                      "
                                                                      class="CToWUd"
                                                                      data-bit="iit"
                                                                    />
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
            
                                                          <table
                                                            align="center"
                                                            border="0"
                                                            cellspacing="0"
                                                            cellpadding="0"
                                                            width="25"
                                                            height="25"
                                                            style="
                                                              width: 25px !important;
                                                              height: 25px !important;
                                                              display: inline-block;
                                                              border-collapse: collapse;
                                                              table-layout: fixed;
                                                              border-spacing: 0;
                                                              vertical-align: top;
                                                              margin-right: 23px;
                                                            "
                                                          >
                                                            <tbody>
                                                              <tr
                                                                style="vertical-align: top"
                                                              >
                                                                <td
                                                                  align="center"
                                                                  valign="middle"
                                                                  style="
                                                                    word-break: break-word;
                                                                    border-collapse: collapse !important;
                                                                    vertical-align: top;
                                                                  "
                                                                >
                                                                  <a
                                                                    href="https://itptit.com/"
                                                                    title="RSS"
                                                                    target="_blank"
                                                                    data-saferedirecturl="https://www.google.com/url?q=https://itptit.com/&amp;source=gmail&amp;ust=1731701344035000&amp;usg=AOvVaw1oePlLFsac9Fvx85OMa9d0"
                                                                  >
                                                                    <img
                                                                      src="https://ci3.googleusercontent.com/meips/ADKq_NZ3XYLwMpWWLMBQwGmDxnPNNqV1OvPr_PfwUQKj7w_j-Bi8G1uLXY7vzJGbM--3M491tauR0seRbCWthJDx9Agia5U8ht2ezRtBhjkIO5Wc=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/rss.png"
                                                                      alt="RSS"
                                                                      title="RSS"
                                                                      width="25"
                                                                      style="
                                                                        outline: none;
                                                                        text-decoration: none;
                                                                        clear: both;
                                                                        display: block !important;
                                                                        border: none;
                                                                        height: auto;
                                                                        float: none;
                                                                        max-width: 25px !important;
                                                                      "
                                                                      class="CToWUd"
                                                                      data-bit="iit"
                                                                    />
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
            
                                                          <table
                                                            align="center"
                                                            border="0"
                                                            cellspacing="0"
                                                            cellpadding="0"
                                                            width="25"
                                                            height="25"
                                                            style="
                                                              width: 25px !important;
                                                              height: 25px !important;
                                                              display: inline-block;
                                                              border-collapse: collapse;
                                                              table-layout: fixed;
                                                              border-spacing: 0;
                                                              vertical-align: top;
                                                              margin-right: 0px;
                                                            "
                                                          >
                                                            <tbody>
                                                              <tr
                                                                style="vertical-align: top"
                                                              >
                                                                <td
                                                                  align="center"
                                                                  valign="middle"
                                                                  style="
                                                                    word-break: break-word;
                                                                    border-collapse: collapse !important;
                                                                    vertical-align: top;
                                                                  "
                                                                >
                                                                  <a
                                                                    href="https://www.tiktok.com/@itclubptithn"
                                                                    title="TikTok"
                                                                    target="_blank"
                                                                    data-saferedirecturl="https://www.google.com/url?q=https://www.tiktok.com/@itclubptithn&amp;source=gmail&amp;ust=1731701344035000&amp;usg=AOvVaw3RYYZ95n6KrSwv8HpfeneG"
                                                                  >
                                                                    <img
                                                                      src="https://ci3.googleusercontent.com/meips/ADKq_NaLjjxiDQ3q_x-6sJTxBD05lKZyuu4RnvlURDp4LnnIH8_-Rr7QjS76BOoJwsCMzMQ7U51QcQ1Gi8taGnmJ1y0L98-YYNVkKfqwd33YDDB_Yw0A=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/tiktok.png"
                                                                      alt="TikTok"
                                                                      title="TikTok"
                                                                      width="25"
                                                                      style="
                                                                        outline: none;
                                                                        text-decoration: none;
                                                                        clear: both;
                                                                        display: block !important;
                                                                        border: none;
                                                                        height: auto;
                                                                        float: none;
                                                                        max-width: 25px !important;
                                                                      "
                                                                      class="CToWUd"
                                                                      data-bit="iit"
                                                                    />
                                                                  </a>
                                                                </td>
                                                              </tr>
                                                            </tbody>
                                                          </table>
                                                        </div>
                                                      </div>
                                                    </td>
                                                  </tr>
                                                </tbody>
                                              </table>
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            </tbody>
                          </table>
                          <div class="yj6qo"></div>
                          <div class="adL"></div>
                        </div>
                      </div>
                      <div class="adL"></div>
                    </div>
                  </div>
                  <div class="WhmR8e" data-hash="0"></div>
                </div>
              </body>
            </html>
        """

        msg.attach(MIMEText(html_body, "html"))

        # Gửi email
        server.sendmail(GMAIL_USER, email, msg.as_string())
        successful_recipients.append({"Name": name, "Email": email})

# Ngắt kết nối với máy chủ
server.quit()

print("Emails sent successfully!")
total_recipients = len(successful_recipients)

# Ghi danh sách người nhận và số lượng vào file
output_file = "E:/JetbrainsTool/PycharmProjects/CodePtit/SinhNhat11/doanhnghiep.txt"

with open(output_file, "w", encoding="utf-8") as f:
    f.write(f"Tổng số người nhận email: {total_recipients}\n\n")
    f.write("Danh sách người nhận:\n")
    for recipient in successful_recipients:
        f.write(f"{recipient['Name']} - {recipient['Email']}\n")

print(f"Emails sent successfully! Total: {total_recipients}")
