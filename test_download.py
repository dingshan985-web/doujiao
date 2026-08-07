import urllib.request as urlrequest
import zipfile
import os

url = 'https://dingshan985-web.github.io/doujiao/update.zip'
print('正在下载测试...')
req = urlrequest.Request(url, headers={'User-Agent': 'doujiao-updater'})
with urlrequest.urlopen(req, timeout=120) as resp:
    print('状态码:', resp.status)
    print('Content-Type:', resp.headers.get('Content-Type'))
    print('Content-Length:', resp.headers.get('Content-Length'))
    data = resp.read()
    print('实际下载大小:', len(data), '字节')
    print('前20字节:', data[:20])
    if data[:2] == b'PK':
        print('✅ 是有效的 ZIP 文件')
        tmp = 'test_download.zip'
        with open(tmp, 'wb') as f:
            f.write(data)
        with zipfile.ZipFile(tmp, 'r') as zf:
            print('ZIP 包含', len(zf.namelist()), '个文件')
            print('前5个文件:', zf.namelist()[:5])
        os.remove(tmp)
    else:
        print('❌ 不是 ZIP 文件')
        print('文本内容前500字:', data[:500].decode('utf-8', errors='replace'))
