"""
================================================================================
 STEP 4 - AI PROCESSING  ->  shared/ai_processor.py
================================================================================

WHAT THIS DOES
--------------
`processAI` is the function described in the requirements. It:
  1. accepts the source document content (bytes) read from SharePoint,
  2. runs it through an AI model,
  3. returns the generated output document (bytes), ready to be uploaded back.

WHERE THE AI ACTUALLY HAPPENS
-----------------------------
The real model call lives in the private helper `_run_model`. The example shown
uses Azure OpenAI, but it is GUARDED: if you have not configured the AI app
settings yet, it skips the network call and returns a clearly-labelled
placeholder. That means you can run and test the WHOLE pipeline end-to-end before
wiring up a model. When you are ready, just set AIEndpoint/AIApiKey/AIDeployment
in app settings, or replace `_run_model` with your own implementation.

INPUT/OUTPUT ARE BYTES
----------------------
SharePoint gives us bytes and expects bytes, so processAI works in bytes too.
Internally we decode to text for the model, then encode the result back.
================================================================================
"""

import logging

from shared.config import Config


def processAI(document_content: bytes) -> bytes:  # noqa: N802 - name fixed by the spec
    """Process the document with AI and return the generated output document.

    Args:
        document_content: raw bytes of the source document (the file located at
            the meeting's `AdminUpdatedFolderLink`).

    Returns:
        The generated output document as bytes.

    Raises:
        ValueError: if no content was supplied.
    """
    if not document_content:
        raise ValueError("processAI received empty document content.")

    # Decode bytes -> text so we can send it to a language model.
    text = _decode(document_content)
    logging.info("processAI: received %d characters of input.", len(text))

    # The actual AI work.
    output_text = _run_model(text)

    logging.info("processAI: produced %d characters of output.", len(output_text))

    # Encode text -> bytes so it can be uploaded back to SharePoint.
    return output_text.encode("utf-8")


def _decode(content: bytes) -> str:
    """Best-effort decode of bytes to a string.

    Most documents are UTF-8. If they are not, we fall back to latin-1 with
    replacement so a stray byte never crashes the whole run.
    """
    try:
        return content.decode("utf-8")
    except UnicodeDecodeError:
        return content.decode("latin-1", errors="replace")


def _run_model(text: str) -> str:
    """Send `text` to the AI model and return its response text.

    REPLACE THIS with your real implementation if you use a different service.
    The reference below is Azure OpenAI.
    """
    # If AI is not configured yet, return a placeholder instead of failing.
    if not (Config.AI_ENDPOINT and Config.AI_API_KEY and Config.AI_DEPLOYMENT):
        logging.warning(
            "AI settings not configured (AIEndpoint/AIApiKey/AIDeployment). "
            "Returning a placeholder document so the pipeline can still run."
        )
        return f"[AI PLACEHOLDER - configure AI settings to enable]\n\n{text}"

    # --- Reference Azure OpenAI call ----------------------------------- #
    # `openai` is installed via requirements.txt.
    from openai import AzureOpenAI

    # Create the client using your Azure OpenAI resource details.
    client = AzureOpenAI(
        azure_endpoint=Config.AI_ENDPOINT,
        api_key=Config.AI_API_KEY,
        api_version="2024-06-01",  # Azure OpenAI API version string
    )

    # A chat completion: a "system" message sets the assistant's job, the
    # "user" message is the transcript text we want processed.
    response = client.chat.completions.create(
        model=Config.AI_DEPLOYMENT,  # the deployment NAME you created in Azure
        messages=[
            {
                "role": "system",
                "content": (
                    "You summarise Microsoft Teams meeting transcripts into a "
                    "clear, structured document with key points, decisions and "
                    "action items."
                ),
            },
            {"role": "user", "content": text},
        ],
        temperature=0.2,  # lower = more focused/consistent, higher = more creative
    )

    # The model's reply text. `or ""` guards against a rare empty response.
    return response.choices[0].message.content or ""
