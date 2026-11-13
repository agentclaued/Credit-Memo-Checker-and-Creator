import os
import tempfile

os.environ["CREDIT_DB_PATH"] = os.path.join(tempfile.mkdtemp(), "test.sqlite")

from fastapi.testclient import TestClient  # noqa: E402

from app.main import app  # noqa: E402


def test_pages_and_lookups():
    with TestClient(app) as c:
        for p in ("/", "/upload", "/memo", "/recheck", "/rules", "/api/health"):
            assert c.get(p).status_code == 200, p
        assert any(i["item"] == "2308ENL" for i in c.get("/api/price-list").json())
        assert any(r["code"] == "DR" for r in c.get("/api/reasons").json())


def test_stylesheets_load_over_http_and_https():
    import re

    for scheme in ("http", "https"):
        with TestClient(app, base_url=f"{scheme}://testserver") as c:
            for page in ("/recheck", "/upload"):
                response = c.get(page)
                links = re.findall(r'<link rel="stylesheet" href="([^"]+)"', response.text)
                local_links = [link for link in links if link.startswith("/static/")]
                assert local_links
                for link in local_links:
                    assert c.get(link).status_code == 200, link


def test_upload_rejects_wrong_type():
    with TestClient(app) as c:
        r = c.post("/upload/ageing", files={"file": ("x.txt", b"hi")})
        assert r.status_code == 400


def test_upload_saves_file_and_lists_it(tmp_path, monkeypatch):
    import app.main as m
    monkeypatch.setattr(m, "UPLOAD_DIR", tmp_path)
    with TestClient(app) as c:
        r = c.post("/upload/orders", files={"file": ("LW9301 test.csv", b"a,b\n1,2\n")})
        assert r.status_code == 200 and r.json()["filename"] == "LW9301 test.csv"
        rows = [f for f in c.get("/api/files").json() if f["filename"] == "LW9301 test.csv"]
        assert rows and rows[0]["kind"] == "orders" and rows[0]["size_bytes"] == 8
        assert list(tmp_path.glob("orders/*LW9301 test.csv"))


def test_memo_script_has_valid_syntax():
    import pathlib, shutil, subprocess
    node = shutil.which("node")
    if not node:
        import pytest
        pytest.skip("node is not installed")
    js = pathlib.Path(__file__).parent.parent / "app" / "static" / "memo.js"
    assert subprocess.run([node, "--check", str(js)], capture_output=True).returncode == 0
