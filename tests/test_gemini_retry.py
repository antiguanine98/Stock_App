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
    assert app._is_retryable_error(RuntimeError("The model is overloaded"))
    assert app._is_retryable_error(RuntimeError("503 UNAVAILABLE"))
    assert not app._is_retryable_error(RuntimeError("invalid argument 400"))


def test_generate_gemini_report_failsover_to_next_model(monkeypatch=None):
    calls: list[str] = []

    class FakeModels:
        def generate_content(self, model, contents, config=None):
            calls.append(model)
            if model == "gemini-3.6-flash":
                raise RuntimeError("503 UNAVAILABLE: model is overloaded")

            class R:
                text = "ok-from-" + model
                candidates = [1]

            return R()

    class FakeClient:
        models = FakeModels()

    app._resolved_gemini_model = "gemini-3.6-flash"
    app.clear_cancel_gemini()

    orig_create = app.create_gemini_client
    orig_resolve = app.resolve_gemini_model
    orig_discover = app._discover_flash_models

    def fake_create(_key):
        return FakeClient()

    def fake_resolve(client, force_refresh=False):
        return "gemini-3.6-flash"

    def fake_discover(client, use_cache: bool = True):
        return ["gemini-3.6-flash", "gemini-omni-1.1-flash"]

    sleeps: list[float] = []

    def fake_sleep(sec):
        sleeps.append(sec)
        return None

    app.create_gemini_client = fake_create
    app.resolve_gemini_model = fake_resolve
    app._discover_flash_models = fake_discover
    orig_sleep = app._sleep_with_cancel
    app._sleep_with_cancel = fake_sleep
    old_max = app._MAX_RETRIES
    app._MAX_RETRIES = 0
    try:
        text = app.generate_gemini_report("fake-key", "hello")
        assert text.startswith("ok-from-")
        assert "gemini-3.6-flash" in calls
        assert any(m != "gemini-3.6-flash" for m in calls)
        assert "gemini-2.5-flash" not in calls
        assert "gemini-1.5-flash" not in calls
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


def test_cascade_models_prefers_discovered_flash():
    models = app._cascade_models(
        "gemini-3.6-flash",
        ["gemini-3.6-flash", "gemini-omni-1.1-flash", "gemini-3.5-flash"],
    )
    assert models[0] == "gemini-3.6-flash"
    assert len(models) <= app._MAX_CASCADE_MODELS
    assert "gemini-2.5-flash" not in models
    assert "gemini-1.5-flash" not in models


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

    orig_disc = app._discover_flash_models
    app._discover_flash_models = lambda c, use_cache=True: [
        "gemini-3.6-flash",
        "gemini-omni-1.1-flash",
    ]
    app._resolved_gemini_model = None
    app._discovered_flash_cache = []
    try:
        model = app.resolve_gemini_model(FakeClient(), force_refresh=True)
        assert model in ("gemini-3.6-flash", "gemini-omni-1.1-flash")
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

    assert app._probe_model(FakeClient(), "gemini-3.6-flash") == "overloaded"


def test_retired_legacy_flash_filtered():
    assert app._is_retired_model("gemini-1.5-flash")
    assert app._is_retired_model("gemini-2.0-flash")
    assert app._is_retired_model("gemini-2.5-flash")
    assert app._is_retired_model("gemini-2.5-flash-lite")
    assert not app._is_retired_model("gemini-3.6-flash")
    assert not app._is_retired_model("gemini-omni-1.1-flash")
    models = app._cascade_models(
        "gemini-2.5-flash",
        ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-3.6-flash"],
    )
    assert "gemini-2.5-flash" not in models
    assert "gemini-1.5-flash" not in models
    assert models[0] == "gemini-3.6-flash"


def test_discover_filters_generate_content_and_orders():
    class FakeModel:
        def __init__(self, name, actions):
            self.name = name
            self.supported_actions = actions

    class FakePager:
        def __iter__(self):
            return iter(
                [
                    FakeModel("models/gemini-2.5-flash", ["generateContent"]),
                    FakeModel("models/gemini-3.6-flash", ["generateContent"]),
                    FakeModel("models/gemini-embedding-flash", ["embedContent"]),
                    FakeModel("models/gemini-omni-1.1-flash", ["generateContent"]),
                    FakeModel("models/gemini-3.1-flash-lite", ["generateContent"]),
                ]
            )

    class FakeClient:
        class models:
            @staticmethod
            def list():
                return FakePager()

    app._discovered_flash_cache = None
    found = app._discover_flash_models(FakeClient(), use_cache=False)
    assert "gemini-2.5-flash" not in found
    assert "gemini-embedding-flash" not in found
    assert found[0] == "gemini-3.6-flash"
    assert "gemini-omni-1.1-flash" in found
    assert "gemini-3.1-flash-lite" in found


def test_http_timeout_at_least_120s():
    assert app._HTTP_TIMEOUT_MS >= 120_000


def test_preference_list_uses_current_server_ids():
    assert "gemini-3.6-flash" in app.GEMINI_MODEL_PREFERENCES
    assert "gemini-2.5-flash" not in app.GEMINI_MODEL_PREFERENCES
    assert "gemini-1.5-flash" not in app.GEMINI_MODEL_PREFERENCES


def test_404_suggestion_appended_to_cascade():
    calls: list[str] = []

    class FakeModels:
        def generate_content(self, model, contents, config=None):
            calls.append(model)
            if model == "gemini-3.6-flash":
                raise RuntimeError(
                    "404 NOT_FOUND. models/gemini-3.6-flash is no longer available "
                    "to new users. Please update your code to use models/gemini-omni-1.1-flash"
                )

            class R:
                text = "ok-" + model
                candidates = [1]

            return R()

    class FakeClient:
        models = FakeModels()

    app._resolved_gemini_model = "gemini-3.6-flash"
    app.clear_cancel_gemini()
    orig_create = app.create_gemini_client
    orig_discover = app._discover_flash_models
    app.create_gemini_client = lambda _k: FakeClient()
    app._discover_flash_models = lambda c, use_cache=True: ["gemini-3.6-flash"]
    try:
        text = app.generate_gemini_report("fake-key", "hi")
        assert text.startswith("ok-")
        assert "gemini-omni-1.1-flash" in calls
    finally:
        app.create_gemini_client = orig_create
        app._discover_flash_models = orig_discover


if __name__ == "__main__":
    test_retryable_detects_503_and_overload_text()
    print("PASS test_retryable_detects_503_and_overload_text")
    test_generate_gemini_report_failsover_to_next_model()
    print("PASS test_generate_gemini_report_failsover_to_next_model")
    test_cancel_aborts_retry_sleep()
    print("PASS test_cancel_aborts_retry_sleep")
    test_cascade_models_prefers_discovered_flash()
    print("PASS test_cascade_models_prefers_discovered_flash")
    test_api_error_code_from_text()
    print("PASS test_api_error_code_from_text")
    test_resolve_accepts_overloaded_as_connected()
    print("PASS test_resolve_accepts_overloaded_as_connected")
    test_probe_model_returns_overloaded_status()
    print("PASS test_probe_model_returns_overloaded_status")
    test_retired_legacy_flash_filtered()
    print("PASS test_retired_legacy_flash_filtered")
    test_discover_filters_generate_content_and_orders()
    print("PASS test_discover_filters_generate_content_and_orders")
    test_http_timeout_at_least_120s()
    print("PASS test_http_timeout_at_least_120s")
    test_preference_list_uses_current_server_ids()
    print("PASS test_preference_list_uses_current_server_ids")
    test_404_suggestion_appended_to_cascade()
    print("PASS test_404_suggestion_appended_to_cascade")
