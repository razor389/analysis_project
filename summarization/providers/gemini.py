import os
from google import genai
from google.genai import types
from .base import LLMProvider

class GeminiProvider(LLMProvider):
    def __init__(self):
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY not found in environment. Please set it in your .env file.")
        self._api_key = api_key

    async def generate_content_async(self, system_prompt, user_prompt, model_name=None, temperature=0.5, max_output_tokens=None):
        model_name = model_name or os.getenv("GEMINI_MODEL", "gemini-2.5-pro")
        
        config = types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="text/plain",
            temperature=temperature,
            max_output_tokens=max_output_tokens,
            # Preserve the existing thresholds for financial/political analysis.
            safety_settings=[
                types.SafetySetting(
                    category=category,
                    threshold=types.HarmBlockThreshold.BLOCK_NONE,
                )
                for category in (
                    types.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
                    types.HarmCategory.HARM_CATEGORY_HARASSMENT,
                    types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
                    types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
                )
            ],
        )

        print(f"[LLM] ==> Calling Gemini ({model_name})...")
        
        try:
            # Each call owns its client, including across separate asyncio.run calls.
            with genai.Client(api_key=self._api_key) as client:
                async with client.aio as async_client:
                    response = await async_client.models.generate_content(
                        model=model_name,
                        contents=[user_prompt],
                        config=config,
                    )
        except Exception as e:
            print(f"[LLM] ==> API Call Failed: {e}")
            return ""

        print("[LLM] ==> Response received.")
        if response.candidates and response.candidates[0].finish_reason == types.FinishReason.MAX_TOKENS:
            print("[LLM] ==> WARNING: MAX_TOKENS. The response might be incomplete.")

        text = (response.text or "").strip()
        if not text:
            print("[LLM] ==> ERROR: Response was empty or blocked.")
            if response.prompt_feedback:
                print(f"[LLM] ==> Prompt Feedback: {response.prompt_feedback}")
            if response.candidates:
                print(f"[LLM] ==> Finish Reason: {response.candidates[0].finish_reason}")
                print(f"[LLM] ==> Safety Ratings: {response.candidates[0].safety_ratings}")
        return text
