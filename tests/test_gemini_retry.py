"""Gemini 과부하 재시도·모델 페일오버 단위 테스트."""

from __future__ import annotations

import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import main as app  # noqa: E402


class _FakeExc(Exception):
    def __init__(self, code: int, message: str = "overloaded"):
        super().__init__(message)
        self.code = code


def test_retryable_detects_503_and_overload_text():
    class E503(Exception):
        pass

    # text-based
    assert app._is_retryable_error(RuntimeError("The model is overloaded"))
    assert app._is_retryable_error(RuntimeError("503 UNAVAILABLE"))
    assert not app._is_retryable_error(RuntimeError("invalid argument 400"))


def test_generate_gemini_report_failsover_to_next_model(monkeypatch=None):
    calls: list[str] = []

    class FakeModels:
        def generate_content(self, model, contents, config=None):
            calls.append(model)
            if model == "gemini-2.5-flash":
                raise RuntimeError("503 UNAVAILABLE: model is overloaded")
            class R:
                text = "ok-from-" + model
                candidates = [1]
            return R()

    class FakeClient:
        models = FakeModels()

    app._resolved_gemini_model = "gemini-2.5-flash"
    app.clear_cancel_gemini()

    # stub helpers
    orig_create = app.create_gemini_client
    orig_resolve = app.resolve_gemini_model
    orig_discover = app._discover_flash_models
    orig_retry = app._call_with_retry

    def fake_create(_key):
        return FakeClient()

    def fake_resolve(client, force_refresh=False):
        return "gemini-2.5-flash"

    def fake_discover(client, use_cache: bool = True):
        return ["gemini-2.5-flash-lite"]

    # Use real _call_with_retry but skip sleep
    sleeps: list[float] = []

    def fake_sleep(sec):
        sleeps.append(sec)
        return None

    app.create_gemini_client = fake_create
    app.resolve_gemini_model = fake_resolve
    app._discover_flash_models = fake_discover
    orig_sleep = app._sleep_with_cancel
    app._sleep_with_cancel = fake_sleep
    # Reduce retries for speed
    old_max = app._MAX_RETRIES
    app._MAX_RETRIES = 0
    try:
        text = app.generate_gemini_report("fake-key", "hello")
        assert text.startswith("ok-from-")
        assert "gemini-2.5-flash" in calls
        # should have moved to another model
        assert any(m != "gemini-2.5-flash" for m in calls)
    finally:
        app._MAX_RETRIES = old_max
        app.create_gemini_client = orig_create
        app.resolve_gemini_model = orig_resolve
        app._discover_flash_models = orig_discover
        app._sleep_with_cancel = orig_sleep


def test_cancel_aborts_retry_sleep():
    app.clear_cancel_gemini()
    app.request_cancel_gemini()
    try:
        app._sleep_with_cancel(5)
        raised = False
    except RuntimeError as exc:
        raised = "취소" in str(exc)
    app.clear_cancel_gemini()
    assert raised




def test_cascade_models_prefers_lite_and_caps():
    models = app._cascade_models("gemini-2.5-flash", ["gemini-2.5-flash", "gemini-2.5-flash-lite"])
    assert models[0] == "gemini-2.5-flash"
    assert len(models) <= app._MAX_CASCADE_MODELS
    # backup은 lite/경량 계열
    assert any("lite" in m or m != "gemini-2.5-flash" for m in models[1:])


def test_api_error_code_from_text():
    assert app._api_error_code(RuntimeError("503 UNAVAILABLE overloaded")) == 503
    assert app._api_error_code(RuntimeError("RESOURCE_EXHAUSTED rate limit")) == 429




def test_resolve_accepts_overloaded_as_connected():
    """모든 후보가 503이어도 API Key 유효 시 모델을 채택한다."""
    class FakeModels:
        def generate_content(self, model, contents, config=None):
            raise RuntimeError("503 UNAVAILABLE: model is overloaded")

    class FakeClient:
        models = FakeModels()

        class models_list:
            pass

    client = FakeClient()
    # discover returns empty via patch
    orig_disc = app._discover_flash_models
    app._discover_flash_models = lambda c, use_cache=True: []
    app._resolved_gemini_model = None
    app._discovered_flash_cache = []
    try:
        model = app.resolve_gemini_model(client, force_refresh=True)
        assert model in app.GEMINI_MODEL_PREFERENCES
        assert app._resolved_gemini_model == model
    finally:
        app._discover_flash_models = orig_disc
        app.clear_cancel_gemini()


def test_probe_model_returns_overloaded_status():
    class FakeModels:
        def generate_content(self, model, contents, config=None):
            raise RuntimeError("503 UNAVAILABLE overloaded")

    class FakeClient:
        models = FakeModels()

    assert app._probe_model(FakeClient(), "gemini-2.0-flash") == "overloaded"




def test_retired_models_filtered_from_cascade():
    assert app._is_retired_model("gemini-2.0-flash")
    assert app._is_retired_model("gemini-2.0-flash-lite")
    models = app._cascade_models("gemini-2.0-flash", ["gemini-2.0-flash", "gemini-3.5-flash-lite"])
    assert "gemini-2.0-flash" not in models
    assert "gemini-3.5-flash-lite" in models


if __name__ == "__main__":
    test_retryable_detects_503_and_overload_text()
    print("PASS test_retryable_detects_503_and_overload_text")
    test_generate_gemini_report_failsover_to_next_model()
    print("PASS test_generate_gemini_report_failsover_to_next_model")
    test_cancel_aborts_retry_sleep()
    print("PASS test_cancel_aborts_retry_sleep")
    test_cascade_models_prefers_lite_and_caps()
    print("PASS test_cascade_models_prefers_lite_and_caps")
    test_api_error_code_from_text()
    print("PASS test_api_error_code_from_text")
    test_resolve_accepts_overloaded_as_connected()
    print("PASS test_resolve_accepts_overloaded_as_connected")
    test_probe_model_returns_overloaded_status()
    print("PASS test_probe_model_returns_overloaded_status")
    test_retired_models_filtered_from_cascade()
    print("PASS test_retired_models_filtered_from_cascade")
