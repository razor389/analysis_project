from abc import ABC, abstractmethod
from typing import Optional

class LLMProvider(ABC):
    @abstractmethod
    async def generate_content_async(
        self,
        system_prompt: str,
        user_prompt: str,
        model_name: Optional[str] = None,
        temperature: float = 0.5,
        max_output_tokens: Optional[int] = None,
    ) -> str:
        """
        Standard interface for generating content.
        Must return the raw text string.
        """
        pass