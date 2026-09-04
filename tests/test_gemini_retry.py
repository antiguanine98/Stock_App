"""Gemini 재시도·모델 고정·Token Diet 단위 테스트."""

from __future__ import annotations

import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import main as app  # noqa: E402


def test_retryable_detects_503_and_overload_text():
    assert app._is_retryable_error(RuntimeError("The model is overloaded"))
    assert app._is_retryable_error(RuntimeError("503 UNAVAILABLE"))
    assert app._is_retryable_error(RuntimeError("429 RESOURCE_EXHAUSTED quota"))
    assert not app._is_retryable_error(RuntimeError("invalid argument 400"))


def test_generate_retries_same_model_not_omni():
    calls: list[str] = []
    sleeps: list[float] = []

    class FakeModels:
        def generate_content(self, model, contents, config=None):
            calls.append(model)
            if len(calls) < 3:
                raise RuntimeError("429 RESOURCE_EXHAUSTED: input_token_count Quota Exceeded")

            class R:
                text = "ok-after-retry"
                candidates = [1]

            return R()

    class FakeClient:
        models = FakeModels()

    app._resolved_gemini_model = "gemini-2.5-flash"
    app.clear_cancel_gemini()
    orig_create = app.create_gemini_client
    orig_sleep = app._sleep_with_cancel
    app.create_gemini_client = lambda _k: FakeClient()
    app._sleep_with_cancel = lambda sec: sleeps.append(sec)
    try:
        text = app.generate_gemini_report("fake-key", "hello")
        assert text == "ok-after-retry"
        assert calls == ["gemini-2.5-flash", "gemini-2.5-flash", "gemini-2.5-flash"]
        assert "omni" not in "".join(calls).lower()
        assert sleeps == [3.0, 3.0]
    finally:
        app.create_gemini_client = orig_create
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


def test_cascade_is_single_fixed_model():
    models = app._cascade_models(
        "gemini-2.5-flash",
        ["gemini-2.5-flash", "gemini-omni-1.1-flash", "gemini-3.6-flash"],
    )
    assert models == ["gemini-2.5-flash"]
    assert app._MAX_CASCADE_MODELS == 1
    assert "omni" not in "".join(models).lower()


def test_omni_and_legacy_retired():
    assert app._is_retired_model("gemini-omni-1.1-flash")
    assert app._is_retired_model("gemini-omni-flash")
    assert app._is_retired_model("gemini-2.0-flash")
    assert not app._is_retired_model("gemini-2.5-flash")
    assert not app._is_retired_model("gemini-1.5-flash-latest")


def test_http_timeout_and_retries():
    assert app._HTTP_TIMEOUT_MS >= 120_000
    assert app._MAX_RETRIES == 2
    assert app._RETRY_BASE_SECONDS == 3
    assert app.GEMINI_MODEL_PREFERENCES == ["gemini-2.5-flash"]
    assert "omni" not in "".join(app.GEMINI_MODEL_PREFERENCES).lower()
    assert "omni" not in "".join(app.GEMINI_FAILOVER_PREFERENCES).lower()


def test_discover_excludes_omni():
    class FakeModel:
        def __init__(self, name, actions):
            self.name = name
            self.supported_actions = actions

    class FakePager:
        def __iter__(self):
            return iter(
                [
                    FakeModel("models/gemini-omni-1.1-flash", ["generateContent"]),
                    FakeModel("models/gemini-2.5-flash", ["generateContent"]),
                    FakeModel("models/gemini-3.6-flash", ["generateContent"]),
                ]
            )

    class FakeClient:
        class models:
            @staticmethod
            def list():
                return FakePager()

    app._discovered_flash_cache = None
    found = app._discover_flash_models(FakeClient(), use_cache=False)
    assert "gemini-omni-1.1-flash" not in found
    assert "gemini-2.5-flash" in found
    assert found[0] == "gemini-2.5-flash"


def test_quota_error_message_clear():
    class FakeModels:
        def generate_content(self, model, contents, config=None):
            raise RuntimeError("429 input_token_count Quota Exceeded")

    class FakeClient:
        models = FakeModels()

    app._resolved_gemini_model = "gemini-2.5-flash"
    app.clear_cancel_gemini()
    orig_create = app.create_gemini_client
    orig_sleep = app._sleep_with_cancel
    app.create_gemini_client = lambda _k: FakeClient()
    app._sleep_with_cancel = lambda sec: None
    try:
        try:
            app.generate_gemini_report("k", "p")
            assert False, "expected raise"
        except RuntimeError as exc:
            assert "할당량" in str(exc) or "Quota" in str(exc)
    finally:
        app.create_gemini_client = orig_create
        app._sleep_with_cancel = orig_sleep


def test_prompt_token_diet_no_full_dump():
    from stock_logic import (
        build_ai_prompt,
        collect_ai_analysis_flags,
        items_for_ai_analysis,
        load_stock_excel,
    )
    from tests import test_stock_logic as t

    path = t._make_sample_xlsx()
    try:
        _, items = load_stock_excel(path)
        flags = collect_ai_analysis_flags(items)
        prompt = build_ai_prompt(items_for_ai_analysis(items), flags=flags)
        assert "Token Diet" in prompt
        assert "추이:[" not in prompt
        assert "구조화 분석 맵 JSON" not in prompt
        assert len(prompt) < 12000
        assert "gemini-omni" not in prompt.lower()
    finally:
        os.remove(path)


if __name__ == "__main__":
    test_retryable_detects_503_and_overload_text()
    print("PASS test_retryable_detects_503_and_overload_text")
    test_generate_retries_same_model_not_omni()
    print("PASS test_generate_retries_same_model_not_omni")
    test_cancel_aborts_retry_sleep()
    print("PASS test_cancel_aborts_retry_sleep")
    test_cascade_is_single_fixed_model()
    print("PASS test_cascade_is_single_fixed_model")
    test_omni_and_legacy_retired()
    print("PASS test_omni_and_legacy_retired")
    test_http_timeout_and_retries()
    print("PASS test_http_timeout_and_retries")
    test_discover_excludes_omni()
    print("PASS test_discover_excludes_omni")
    test_quota_error_message_clear()
    print("PASS test_quota_error_message_clear")
    test_prompt_token_diet_no_full_dump()
    print("PASS test_prompt_token_diet_no_full_dump")
