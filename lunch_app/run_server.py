"""社内 PC で本番用サーバー（waitress）を起動する。

Windows でも動くよう、開発用サーバーではなく waitress を使う。
同じ社内ネットワークの端末からアクセスするための URL も表示する。
"""
import os

from waitress import serve

from app import create_app
from netinfo import lan_ip

PORT = int(os.environ.get("PORT", "5002"))


def main() -> None:
    app = create_app()
    address = lan_ip()
    line = "=" * 56
    print(line)
    print("  社内昼食注文システムを起動しました")
    print(line)
    print(f"  このPC          : http://localhost:{PORT}")
    print(f"  社内の他の端末  : http://{address}:{PORT}")
    print()
    print("  スマートフォンからは上の「社内の他の端末」のURLを開いてください。")
    print("  終了するには、この画面で Ctrl + C を押すか、ウィンドウを閉じます。")
    print(line)
    serve(app, host="0.0.0.0", port=PORT, threads=8)


if __name__ == "__main__":
    main()
