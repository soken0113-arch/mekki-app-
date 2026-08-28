"""社内 LAN でのアクセス先 URL を組み立てるためのヘルパー。"""
import socket


def lan_ip() -> str:
    """社内 LAN で他の端末から見える、この PC の IP アドレス。

    外部に送信はせず、経路表から自分側のアドレスを取得するだけ。
    取得できない場合は 127.0.0.1 を返す（＝他の端末からは繋がらない状態）。
    """
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("8.8.8.8", 80))
        return sock.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        sock.close()


def access_url(port: str | int = 5002) -> str:
    """社員に伝えるアクセス先 URL。"""
    port = str(port)
    host = lan_ip()
    return f"http://{host}" if port == "80" else f"http://{host}:{port}"
