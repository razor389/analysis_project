"""Gemini SDK regression tests with real serialization and mocked HTTP responses."""

import asyncio
from collections import deque
from contextlib import redirect_stdout
import importlib
import io
import json
import os
import unittest
from unittest.mock import patch

from google import genai
from google.genai import types
import httpx


def text_response(text, finish_reason="STOP", *, thought=False):
    return {
        "candidates": [{
            "content": {
                "role": "model",
                "parts": [{"text": text, "thought": thought}],
            },
            "finishReason": finish_reason,
        }],
    }


class RecordingTransport(httpx.MockTransport):
    closed = False

    def close(self):
        self.closed = True
        super().close()

    async def aclose(self):
        self.closed = True
        await super().aclose()


class GeminiTests(unittest.TestCase):
    def setUp(self):
        self.enterContext(patch.dict(os.environ, {
            "GEMINI_API_KEY": "test-gemini-key",
            "GEMINI_MODEL": "gemini-2.5-pro",
            "GOOGLE_API_KEY": "",
            "LLM_PROVIDER": "gemini",
        }))
        self.enterContext(patch("dotenv.load_dotenv", return_value=False))
        self.core = importlib.import_module("summarization.core")
        self.provider_module = importlib.import_module("summarization.providers.gemini")
        self.standalone = importlib.import_module("gemini_summarizer")
        self.enterContext(patch.object(self.standalone, "gemini_api_key", "test-gemini-key"))
        self.provider = self.provider_module.GeminiProvider()
        self.implementations = (
            self.provider.generate_content_async,
            self.standalone._call_gemini_async,
        )
        self.output = self.enterContext(redirect_stdout(io.StringIO()))
        self.responses = deque()
        self.requests = []
        self.transports = []
        real_client = genai.Client

        async def handle_request(request):
            # Yield so concurrent callers exercise separate client lifetimes.
            await asyncio.sleep(0)
            self.requests.append(request)
            response = self.responses.popleft()
            if isinstance(response, Exception):
                raise response
            return response

        def create_client(**kwargs):
            sync_transport = RecordingTransport(handle_request)
            async_transport = RecordingTransport(handle_request)
            self.transports.extend((sync_transport, async_transport))
            return real_client(
                **kwargs,
                vertexai=False,
                http_options=types.HttpOptions(
                    client_args={"transport": sync_transport, "trust_env": False},
                    async_client_args={"transport": async_transport, "trust_env": False},
                    retry_options=types.HttpRetryOptions(attempts=1),
                ),
            )

        self.enterContext(patch.object(genai, "Client", side_effect=create_client))

    def tearDown(self):
        self.assertFalse(self.responses, "Expected HTTP responses were not consumed")
        for transport in self.transports:
            self.assertTrue(transport.closed, "An SDK HTTP client was left open")

    def respond(self, body, status=200):
        self.responses.append(httpx.Response(status, json=body))

    def test_request_preserves_prompts_and_generation_settings(self):
        for generate in self.implementations:
            with self.subTest(implementation=generate.__qualname__):
                self.respond(text_response("  Summary text.\n"))
                result = asyncio.run(generate(
                    "System instructions", "User posts", model_name="test-model",
                    temperature=0.3, max_output_tokens=8192,
                ))
                self.assertEqual(result, "Summary text.")
                request = self.requests[-1]
                self.assertEqual(request.method, "POST")
                self.assertTrue(request.url.path.endswith("/models/test-model:generateContent"))
                self.assertEqual(request.headers["x-goog-api-key"], "test-gemini-key")
                body = json.loads(request.content)
                self.assertEqual(body["systemInstruction"]["parts"], [{"text": "System instructions"}])
                self.assertEqual(body["contents"], [{"role": "user", "parts": [{"text": "User posts"}]}])
                self.assertEqual(body["generationConfig"], {
                    "temperature": 0.3, "maxOutputTokens": 8192,
                    "responseMimeType": "text/plain",
                })
                self.assertEqual(
                    {setting["category"]: setting["threshold"] for setting in body["safetySettings"]},
                    {
                        "HARM_CATEGORY_HATE_SPEECH": "BLOCK_NONE",
                        "HARM_CATEGORY_HARASSMENT": "BLOCK_NONE",
                        "HARM_CATEGORY_SEXUALLY_EXPLICIT": "BLOCK_NONE",
                        "HARM_CATEGORY_DANGEROUS_CONTENT": "BLOCK_NONE",
                    },
                )

    def test_default_settings(self):
        for generate in self.implementations:
            with self.subTest(implementation=generate.__qualname__):
                self.respond(text_response("Summary"))
                self.assertEqual(asyncio.run(generate("System", "Posts")), "Summary")
                request = self.requests[-1]
                self.assertTrue(request.url.path.endswith("/models/gemini-2.5-pro:generateContent"))
                self.assertEqual(json.loads(request.content)["generationConfig"], {
                    "temperature": 0.5, "responseMimeType": "text/plain",
                })

    def test_provider_uses_environment_model(self):
        self.respond(text_response("Summary"))
        with patch.dict(os.environ, {"GEMINI_MODEL": "configured-model"}):
            asyncio.run(self.provider.generate_content_async("System", "Posts"))
        self.assertTrue(self.requests[-1].url.path.endswith("/models/configured-model:generateContent"))

    def test_standalone_preserves_custom_mime_type(self):
        self.respond(text_response('{"summary": "text"}'))
        asyncio.run(self.standalone._call_gemini_async(
            "System", "Posts", response_mime_type="application/json",
        ))
        self.assertEqual(json.loads(self.requests[-1].content)["generationConfig"]["responseMimeType"], "application/json")

    def test_empty_and_blocked_responses_return_empty_text(self):
        responses = (
            {},
            text_response("   "),
            {"promptFeedback": {"blockReason": "SAFETY"}},
            {"candidates": [{"finishReason": "SAFETY", "safetyRatings": [{
                "category": "HARM_CATEGORY_HARASSMENT", "probability": "HIGH", "blocked": True,
            }]}]},
        )
        for generate in self.implementations:
            for response in responses:
                with self.subTest(implementation=generate.__qualname__, response=response):
                    self.respond(response)
                    self.assertEqual(asyncio.run(generate("System", "Posts")), "")
        self.assertIn("Response was empty or blocked", self.output.getvalue())
        self.assertIn("Prompt Feedback", self.output.getvalue())
        self.assertIn("Safety Ratings", self.output.getvalue())

    def test_max_tokens_preserves_partial_text(self):
        for generate in self.implementations:
            with self.subTest(implementation=generate.__qualname__):
                self.respond(text_response("  Partial summary  ", "MAX_TOKENS"))
                self.assertEqual(asyncio.run(generate("System", "Posts")), "Partial summary")
        self.assertEqual(self.output.getvalue().count("WARNING: MAX_TOKENS"), 2)

    def test_max_tokens_with_only_thoughts_returns_empty_text(self):
        for generate in self.implementations:
            with self.subTest(implementation=generate.__qualname__):
                self.respond(text_response("Internal reasoning", "MAX_TOKENS", thought=True))
                self.assertEqual(asyncio.run(generate("System", "Posts")), "")
        self.assertEqual(self.output.getvalue().count("WARNING: MAX_TOKENS"), 2)

    def test_api_and_connection_errors_return_empty_text(self):
        for generate in self.implementations:
            for failure in (
                httpx.Response(400, json={"error": {"code": 400, "message": "Bad request", "status": "INVALID_ARGUMENT"}}),
                httpx.ConnectError("Connection failed"),
            ):
                with self.subTest(implementation=generate.__qualname__, failure=type(failure).__name__):
                    self.responses.append(failure)
                    self.assertEqual(asyncio.run(generate("System", "Posts")), "")
        self.assertEqual(self.output.getvalue().count("API Call Failed"), 4)

    def test_repeated_event_loops_and_concurrent_requests(self):
        async def concurrent_calls(generate):
            return await asyncio.gather(generate("System", "First"), generate("System", "Second"))

        for generate in self.implementations:
            with self.subTest(implementation=generate.__qualname__):
                for _ in range(2):
                    self.respond(text_response("Summary"))
                    self.assertEqual(asyncio.run(generate("System", "Posts")), "Summary")
                for _ in range(2):
                    self.respond(text_response("Summary"))
                self.assertEqual(asyncio.run(concurrent_calls(generate)), ["Summary", "Summary"])

    def test_public_post_summary_interfaces(self):
        summary = "1. **Margins:** Stable."
        for module in (self.core, self.standalone):
            with self.subTest(module=module.__name__):
                self.respond(text_response("Here is the summary:\n" + summary))
                self.assertEqual(module.generate_post_summary([{"message": "Margins are stable."}], "SBUX"), summary)

    def test_moat_summary_retries_empty_response(self):
        for module in (self.core, self.standalone):
            with self.subTest(module=module.__name__):
                self.respond({})
                self.respond(text_response("Risk summary."))
                result = asyncio.run(module._generate_moat_subcategory_summary_async(
                    [{"message": "Competition is increasing."}], "SBUX", "Competition",
                ))
                self.assertEqual(result, "Risk summary.")
                config = json.loads(self.requests[-1].content)["generationConfig"]
                self.assertEqual(config["temperature"], 0.2)
                self.assertEqual(config["maxOutputTokens"], 8192)

    def test_missing_api_key(self):
        with patch.dict(os.environ, {"GEMINI_API_KEY": ""}):
            with self.assertRaisesRegex(ValueError, "GEMINI_API_KEY not found"):
                self.provider_module.GeminiProvider()


if __name__ == "__main__":
    unittest.main()
