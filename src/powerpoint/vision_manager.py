import os
from google import genai
from google.genai import types
from PIL import Image
from io import BytesIO


class VisionManager:

    async def generate_and_save_image(self, prompt: str,
                                      output_path: str) -> str:
        """Generate an image using Gemini Model and save it to the specified path."""

        api_key = os.environ.get('GEMINI_API_KEY')
        if not api_key:
            raise ValueError("GEMINI_API_KEY environment variable not set.")

        client = genai.Client(api_key=api_key)

        try:
            # Generate the image
            response = client.models.generate_content(
                model="gemini-2.0-flash-preview-image-generation",
                contents=(prompt),
                config=types.GenerateContentConfig(
                    response_modalities=['TEXT', 'IMAGE']))

        except Exception as e:
            raise ValueError(f"Failed to generate image: {str(e)}")

        # Save the image
        try:
            img_data = [
                part.inline_data
                for part in response.candidates[0].content.parts
                if part.inline_data is not None
            ][0].data
            image = Image.open(BytesIO((img_data)))
            # Ensure the save directory exists
            try:
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
            except OSError as e:
                raise ValueError(
                    f"Failed to create a directory for image: str({e})")
            # Save the image
            image.save(output_path)
        except (IOError, OSError) as e:
            raise ValueError(
                f"Failed to save image to {output_path}: {str(e)}")

        return output_path
