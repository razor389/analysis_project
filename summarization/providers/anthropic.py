import os
from .base import LLMProvider

class AnthropicProvider(LLMProvider):
    def __init__(self):
        from anthropic import AsyncAnthropic
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY not found in environment.")
        self.client = AsyncAnthropic(api_key=api_key)

    async def generate_content_async(self, system_prompt, user_prompt, model_name=None, temperature=0.5, max_output_tokens=None):
        model_name = model_name or os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-20241022")
        # Claude defaults to 4096 if not specified
        max_tokens = max_output_tokens or 4096
        
        print(f"[LLM] ==> Calling Anthropic ({model_name})...")
        try:
            message = await self.client.messages.create(
                model=model_name,
                max_tokens=max_tokens,
                temperature=temperature,
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}]
            )
            return message.content[0].text.strip()
        except Exception as e:
            print(f"[LLM] ==> Anthropic Error: {e}")
            return ""