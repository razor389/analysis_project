import os
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold, GenerationConfig
from .base import LLMProvider

class GeminiProvider(LLMProvider):
    def __init__(self):
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY not found in environment. Please set it in your .env file.")
        genai.configure(api_key=api_key)

    async def generate_content_async(self, system_prompt, user_prompt, model_name=None, temperature=0.5, max_output_tokens=None):
        model_name = model_name or os.getenv("GEMINI_MODEL", "gemini-2.5-pro")
        
        model = genai.GenerativeModel(
            model_name=model_name,
            system_instruction=system_prompt,
        )
        
        config = GenerationConfig(
            response_mime_type="text/plain",
            temperature=temperature,
            max_output_tokens=max_output_tokens,
        )

        # BLOCK_NONE is essential for financial/political analysis to avoid false positives
        safety_settings = {
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
        }

        print(f"[LLM] ==> Calling Gemini ({model_name})...")
        
        try:
            response = await model.generate_content_async(
                [user_prompt],
                generation_config=config,
                safety_settings=safety_settings
            )
            print("[LLM] ==> Response received.")
            
            # Check for finish reason 2 (MAX_TOKENS)
            if response.candidates and response.candidates[0].finish_reason == 2:
                print("[LLM] ==> WARNING: Finish Reason 2 (MAX_TOKENS). The response might be incomplete.")
                if hasattr(response, 'text'):
                    return response.text.strip()
                return ""

            return response.text.strip()

        except ValueError:
            print("[LLM] ==> ERROR: Response was empty or blocked.")
            # Attempt to debug specific Gemini block reasons
            try:
                if response.prompt_feedback:
                    print(f"[LLM] ==> Prompt Feedback: {response.prompt_feedback}")
                if response.candidates:
                    print(f"[LLM] ==> Finish Reason: {response.candidates[0].finish_reason}")
                    print(f"[LLM] ==> Safety Ratings: {response.candidates[0].safety_ratings}")
            except:
                pass
            return ""
            
        except Exception as e:
            print(f"[LLM] ==> API Call Failed: {e}")
            return ""