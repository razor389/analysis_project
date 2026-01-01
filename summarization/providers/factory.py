import os
from .base import LLMProvider

def get_provider() -> LLMProvider:
    # Defaults to 'gemini' if not set
    provider_name = os.getenv("LLM_PROVIDER", "gemini").lower().strip()
    
    if provider_name == "anthropic":
        from .anthropic import AnthropicProvider
        return AnthropicProvider()
        
    elif provider_name == "openai":
        from .openai import OpenAIProvider
        return OpenAIProvider()
        
    else:
        from .gemini import GeminiProvider
        return GeminiProvider()