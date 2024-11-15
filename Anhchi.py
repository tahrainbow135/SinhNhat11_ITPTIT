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
df = pd.read_excel("E:/JetbrainsTool/PycharmProjects/CodePtit/SinhNhat11/Anhchi.xlsx")

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
                <div>
                  <div class="adM"></div>
                  <table
                    id="m_1768578558709404887u_body"
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
                                                    background-color: rgb(255, 255, 255);
                                                    padding-bottom: 12pt;
                                                    text-align: justify;
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Thân gửi
                                                    <strong>{name},</strong></span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    background-color: rgb(255, 255, 255);
                                                    padding-bottom: 12pt;
                                                    text-align: justify;
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Lời đầu tiên, chúng em xin gửi đến </span
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
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                  >
                                                    lời chào thân ái cùng những lời chúc tốt
                                                    đẹp nhất.</span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    background-color: rgb(255, 255, 255);
                                                    padding-bottom: 12pt;
                                                    text-align: justify;
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Năm 2024 đánh dấu một chặng đường đầy ý
                                                    nghĩa của đại gia đình</span
                                                  ><span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #ff9900;
                                                      line-height: 28.8px;
                                                    "
                                                  >
                                                  </span
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
                                                  ><strong
                                                    ><span
                                                      style="
                                                        font-size: 14pt;
                                                        font-family: 'Times New Roman',
                                                          serif;
                                                        color: #000000;
                                                        line-height: 28.8px;
                                                      "
                                                      >.
                                                    </span></strong
                                                  ><span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Nhìn lại hành trình 11 năm, tuy không
                                                    dài, nhưng mỗi bước đi đều được ghi dấu
                                                    trong trái tim từng thành viên của
                                                    CLB.&nbsp;</span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    background-color: rgb(255, 255, 255);
                                                    padding-bottom: 12pt;
                                                    text-align: justify;
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Qua bao thời gian, CLB</span
                                                  ><strong
                                                    ><span
                                                      style="
                                                        font-size: 14pt;
                                                        font-family: 'Times New Roman',
                                                          serif;
                                                        color: #e69138;
                                                        line-height: 28.8px;
                                                      "
                                                    >
                                                    </span></strong
                                                  ><span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >chúng ta</span
                                                  ><strong
                                                    ><span
                                                      style="
                                                        font-size: 14pt;
                                                        font-family: 'Times New Roman',
                                                          serif;
                                                        color: #e69138;
                                                        line-height: 28.8px;
                                                      "
                                                    >
                                                    </span></strong
                                                  ><span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >không chỉ mang đến kiến thức mà còn
                                                    truyền cảm hứng và tinh thần gắn kết đến
                                                    bao thế hệ sinh viên PTIT. Để có được
                                                    thành quả như hôm nay, không thể không
                                                    nhắc tới những đóng góp âm thầm nhưng
                                                    đầy trân quý từ các anh chị đi trước –
                                                    những người luôn dốc lòng xây dựng và
                                                    phát triển CLB</span
                                                  ><strong
                                                    ><span
                                                      style="
                                                        font-size: 14pt;
                                                        font-family: 'Times New Roman',
                                                          serif;
                                                        color: #000000;
                                                        line-height: 28.8px;
                                                      "
                                                      >.</span
                                                    ></strong
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    background-color: rgb(255, 255, 255);
                                                    padding-bottom: 12pt;
                                                    text-align: justify;
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Để bày tỏ lòng tri ân sâu sắc đến các
                                                    thế hệ tiền bối và đánh d</span
                                                  ><span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >ấu một cột mốc đáng nhớ, </span
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
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                  >
                                                    trân trọng tổ chức </span
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
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                  >
                                                    thành lập. Đây sẽ là dịp đặc biệt để
                                                    chúng ta cùng nhau ôn lại chặng đường đã
                                                    qua với lòng biết ơn, đồng thời hướng
                                                    tới tương lai với niềm tin và khát vọng
                                                    lớn lao.</span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    background-color: rgb(255, 255, 255);
                                                    padding-bottom: 12pt;
                                                    text-align: justify;
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Xin chân thành cảm ơn </span
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
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                  >
                                                    vì những đồng hành và đóng góp quý báu
                                                    trong suốt hành trình này. Rất mong được
                                                    đón tiếp </span
                                                  ><span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >anh/chị</span
                                                  ><span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                  >
                                                    tại bữa tiệc sinh nhật với chủ đề </span
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
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >.</span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    text-align: justify;
                                                    background-color: rgb(255, 255, 255);
                                                    padding-bottom: 12pt;
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Thông tin chi tiết về buổi lễ kỷ
                                                    niệm:</span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    text-align: justify;
                                                    background-color: rgb(255, 255, 255);
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
                                                      >Thời gian: 17h30, Thứ Bảy, ngày 23
                                                      tháng 11 năm 2024</span
                                                    ></strong
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    text-align: justify;
                                                    background-color: rgb(255, 255, 255);
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
                                                      >Địa điểm: Hội trường A2, Học viện
                                                      Công nghệ Bưu chính Viễn thông (</span
                                                    ></strong
                                                  ><a
                                                    href="https://www.google.com/maps/dir/20.9752565,105.7852988/20.9807365,105.7874288/@20.9780266,105.7800735,16z/data=!3m1!4b1!4m4!4m3!1m1!4e1!1m0?entry=ttu&amp;g_ep=EgoyMDI0MTExMS4wIKXMDSoASAFQAw%3D%3D"
                                                    style="text-decoration: none"
                                                    target="_blank"
                                                    data-saferedirecturl="https://www.google.com/url?q=https://www.google.com/maps/dir/20.9752565,105.7852988/20.9807365,105.7874288/@20.9780266,105.7800735,16z/data%3D!3m1!4b1!4m4!4m3!1m1!4e1!1m0?entry%3Dttu%26g_ep%3DEgoyMDI0MTExMS4wIKXMDSoASAFQAw%253D%253D&amp;source=gmail&amp;ust=1731704122596000&amp;usg=AOvVaw3fgkKUjGytHqZOhvBuzRO_"
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
                                                        >Km 10 Nguyễn Trãi, Hà Đông, Hanoi,
                                                        Vietnam</span
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
                                                    line-height: 160%;
                                                    text-align: justify;
                                                    background-color: rgb(255, 255, 255);
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Trân trọng,</span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    text-align: justify;
                                                    background-color: rgb(255, 255, 255);
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
                                                      color: #000000;
                                                      line-height: 28.8px;
                                                    "
                                                    >Lê Đức Hiếu</span
                                                  >
                                                </p>
                                                <p
                                                  style="
                                                    line-height: 160%;
                                                    text-align: justify;
                                                    background-color: rgb(255, 255, 255);
                                                    margin: 0px;
                                                  "
                                                >
                                                  <span
                                                    style="
                                                      font-size: 14pt;
                                                      font-family: 'Times New Roman', serif;
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
                                                  style="display: table; max-width: 239px"
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
                                                      <tr style="vertical-align: top">
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
                                                      <tr style="vertical-align: top">
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
                                                            data-saferedirecturl="https://www.google.com/url?q=https://www.facebook.com/ITPTIT&amp;source=gmail&amp;ust=1731704122597000&amp;usg=AOvVaw1CWtW11YKp1sWQ0ywh7lmE"
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
                                                      <tr style="vertical-align: top">
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
                                                            data-saferedirecturl="https://www.google.com/url?q=https://www.youtube.com/channel/UC8Iwsz8PT07_yVpqEvG7MRw&amp;source=gmail&amp;ust=1731704122597000&amp;usg=AOvVaw0q-FHIRx3Q7m741Tv_lWOn"
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
                                                      <tr style="vertical-align: top">
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
                                                            data-saferedirecturl="https://www.google.com/url?q=https://itptit.com/&amp;source=gmail&amp;ust=1731704122597000&amp;usg=AOvVaw3wuZVsmqBJ_B6p_NXeFKGk"
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
                                                      <tr style="vertical-align: top">
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
                                                            data-saferedirecturl="https://www.google.com/url?q=https://www.tiktok.com/@itclubptithn&amp;source=gmail&amp;ust=1731704122597000&amp;usg=AOvVaw0dr7jXzjUTSazb0RZnyLjc"
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
            
                  <br />
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
output_file = "E:/JetbrainsTool/PycharmProjects/CodePtit/SinhNhat11/anhchi.txt"

with open(output_file, "w", encoding="utf-8") as f:
    f.write(f"Tổng số người nhận email: {total_recipients}\n\n")
    f.write("Danh sách người nhận:\n")
    for recipient in successful_recipients:
        f.write(f"{recipient['Name']} - {recipient['Email']}\n")

print(f"Emails sent successfully! Total: {total_recipients}")
