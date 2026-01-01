import os
from .base import LLMProvider

class OpenAIProvider(LLMProvider):
    def __init__(self):
        from openai import AsyncOpenAI
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY not found in environment.")
        self.client = AsyncOpenAI(api_key=api_key)

    async def generate_content_async(self, system_prompt, user_prompt, model_name=None, temperature=0.5, max_output_tokens=None):
        model_name = model_name or os.getenv("OPENAI_MODEL", "gpt-4o")
        # GPT-4o typically supports up to 4096 output tokens in basic tier, sometimes more
        max_tokens = max_output_tokens or 4096
        
        print(f"[LLM] ==> Calling OpenAI ({model_name})...")
        try:
            response = await self.client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=temperature,
                max_tokens=max_tokens,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            print(f"[LLM] ==> OpenAI Error: {e}")
            return ""