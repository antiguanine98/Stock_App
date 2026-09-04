"""Gemini 연결 탐색·재시도·Token Diet 단위 테스트."""

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

    app._resolved_gemini_model = "gemini-3.6-flash"
    app.clear_cancel_gemini()
    orig_create = app.create_gemini_client
    orig_sleep = app._sleep_with_cancel
    app.create_gemini_client = lambda _k: FakeClient()
    app._sleep_with_cancel = lambda sec: sleeps.append(sec)
    try:
        text = app.generate_gemini_report("fake-key", "hello")
        assert text == "ok-after-retry"
        assert all(c == "gemini-3.6-flash" for c in calls)
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


def test_default_model_is_36_flash():
    assert app.GEMINI_DEFAULT_MODEL == "gemini-3.6-flash"
    assert app.GEMINI_MODEL_PREFERENCES[0] == "gemini-3.6-flash"
    assert "omni" not in "".join(app.GEMINI_MODEL_PREFERENCES).lower()
    assert app._HTTP_TIMEOUT_MS >= 120_000


def test_omni_retired_36_allowed():
    assert app._is_retired_model("gemini-omni-1.1-flash")
    assert not app._is_retired_model("gemini-3.6-flash")
    assert not app._is_retired_model("gemini-2.5-flash")


def test_discover_prefers_36_flash_case_insensitive():
    class FakeModel:
        def __init__(self, name, actions=None, methods=None):
            self.name = name
            self.supported_actions = actions
            self.supported_generation_methods = methods

    class FakePager:
        def __iter__(self):
            return iter(
                [
                    FakeModel("models/GEMINI-2.5-FLASH", methods=["generateContent"]),
                    FakeModel("models/gemini-omni-1.1-flash", actions=["generateContent"]),
                    FakeModel("models/Gemini-3.6-Flash", actions=["generateContent"]),
                    FakeModel("models/gemini-pro", methods=["generateContent"]),
                ]
            )

    class FakeClient:
        class models:
            @staticmethod
            def list():
                return FakePager()

    app._discovered_flash_cache = None
    found = app._discover_flash_models(FakeClient(), use_cache=False)
    assert found[0] == "Gemini-3.6-Flash" or "3.6-flash" in found[0].lower()
    assert "omni" not in "".join(found).lower()


def test_discover_empty_falls_back_to_default():
    class FakeClient:
        class models:
            @staticmethod
            def list():
                raise RuntimeError("network")

    app._discovered_flash_cache = None
    found = app._discover_flash_models(FakeClient(), use_cache=False)
    assert found[0] == app.GEMINI_DEFAULT_MODEL
    assert app.GEMINI_DEFAULT_MODEL in found


def test_resolve_never_raises_model_not_found():
    """핑이 전부 실패해도 기본 모델로 연결 성공."""

    class FakeModels:
        def generate_content(self, model, contents, config=None):
            raise RuntimeError("404 NOT_FOUND")

        def list(self):
            return []

    class FakeClient:
        models = FakeModels()

    app._resolved_gemini_model = None
    app._discovered_flash_cache = None
    model = app.resolve_gemini_model(FakeClient(), force_refresh=True)
    assert model == app.GEMINI_DEFAULT_MODEL
    assert app._resolved_gemini_model == model


def test_connection_returns_connected_message():
    class FakeModels:
        def generate_content(self, model, contents, config=None):
            class R:
                text = "pong"
                candidates = [1]

            return R()

        def list(self):
            class M:
                name = "models/gemini-3.6-flash"
                supported_actions = ["generateContent"]

            return [M()]

    class FakeClient:
        models = FakeModels()

    orig = app.create_gemini_client
    app.create_gemini_client = lambda _k: FakeClient()
    app._resolved_gemini_model = None
    app._discovered_flash_cache = None
    try:
        ok, msg = app.test_gemini_connection("fake-key")
        assert ok is True
        assert msg.startswith("연결됨")
        assert "gemini-3.6-flash" in msg.lower() or "3.6-flash" in msg.lower()
    finally:
        app.create_gemini_client = orig


def test_pick_target_priority():
    assert (
        app._pick_target_from_available(
            ["gemini-2.5-flash", "gemini-3.6-flash", "gemini-pro"]
        )
        == "gemini-3.6-flash"
    )
    assert (
        app._pick_target_from_available(["gemini-pro", "other-flash-x"])
        == "other-flash-x"
    )
    assert app._pick_target_from_available(["gemini-pro"]) == "gemini-pro"
    assert app._pick_target_from_available([]) is None


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
        assert len(prompt) < 12000
    finally:
        os.remove(path)


if __name__ == "__main__":
    test_retryable_detects_503_and_overload_text()
    print("PASS test_retryable_detects_503_and_overload_text")
    test_generate_retries_same_model_not_omni()
    print("PASS test_generate_retries_same_model_not_omni")
    test_cancel_aborts_retry_sleep()
    print("PASS test_cancel_aborts_retry_sleep")
    test_default_model_is_36_flash()
    print("PASS test_default_model_is_36_flash")
    test_omni_retired_36_allowed()
    print("PASS test_omni_retired_36_allowed")
    test_discover_prefers_36_flash_case_insensitive()
    print("PASS test_discover_prefers_36_flash_case_insensitive")
    test_discover_empty_falls_back_to_default()
    print("PASS test_discover_empty_falls_back_to_default")
    test_resolve_never_raises_model_not_found()
    print("PASS test_resolve_never_raises_model_not_found")
    test_connection_returns_connected_message()
    print("PASS test_connection_returns_connected_message")
    test_pick_target_priority()
    print("PASS test_pick_target_priority")
    test_prompt_token_diet_no_full_dump()
    print("PASS test_prompt_token_diet_no_full_dump")
