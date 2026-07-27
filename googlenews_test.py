from collector import gnewsdecoder
google_news_url = "https://news.google.com/rss/articles/CBMihgFBVV95cUxQOW9qOC11c2lYUGdGMWE5N1hwRW5WV21lMndKczhEc21LdzJfRlhSRGJhUUY5ZWdicG9wakpUVEFPTEdqM0FQRWRvekc2c2lidmxUWXBDdFlFTy1ZLVYxeFJLSzdTLWxhM25Nbm5YU0ZPYWp4ZWU5RFctZ1Z5M1ZNcUIwSXBGZw?oc=5"

result = gnewsdecoder(google_news_url)
print(result)
if result.get("status"):
    print(result["decoded_url"])
else:
    print("변환 실패:", result.get("message"))