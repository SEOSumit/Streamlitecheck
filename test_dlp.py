import json
from dlp_content_audit import process_single_url

def main():
    result = process_single_url("https://www.lenovo.com/jp/ja/d/15-inch-laptops/")
    print(json.dumps(result, indent=2))

if __name__ == "__main__":
    main()
