from iso_crawler import ISO_Crawler

if __name__ == "__main__":
    start = True
    while start:
        url = input("Enter the website URL(請輸入網站url): ")
        if not url:
            print("Please enter a valid URL(請輸入有效的網站url)")
            continue
        else:
            start = False
            output_path = input("Enter the output file path (輸入欲輸出文件名稱，default(預設值) : 'crawled_content.docx'): ")
            if not output_path:
                output_path = "crawled_content.docx"
            ISO_Crawler(url, output_path)