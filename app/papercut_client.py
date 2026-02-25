import ssl
import xmlrpc.client
from typing import Any, Optional


class PaperCutClient:
    def __init__(self, url: str, auth_token: str, verify_tls: bool = True) -> None:
        self.url = url
        self.auth_token = auth_token
        self.verify_tls = verify_tls

        context = None
        if url.lower().startswith("https") and not verify_tls:
            context = ssl._create_unverified_context()

        self._proxy = xmlrpc.client.ServerProxy(url, allow_none=True, context=context)

    def call(self, method: str, *params: Any) -> Any:
        api = self._proxy.api
        fn = getattr(api, method)
        return fn(self.auth_token, *params)


def build_client(url: Optional[str], token: Optional[str], verify_tls: bool) -> Optional[PaperCutClient]:
    if not url or not token:
        return None
    return PaperCutClient(url, token, verify_tls)
