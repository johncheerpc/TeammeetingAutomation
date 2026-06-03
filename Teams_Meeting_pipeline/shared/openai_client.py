# =============================================================================
# FILE: shared/openai_client.py
#
# PURPOSE:
#   Sends the Teams meeting transcript (and reference document text) to
#   Azure OpenAI and gets back a structured JSON analysis of the meeting.
#
# WHAT IS AZURE OPENAI?
#   Azure OpenAI is Microsoft's hosted version of GPT-4o (and other models).
#   Instead of using OpenAI's public API, we use the Azure-hosted version
#   because it stays within our Azure tenant (better security & compliance).
#
# AUTHENTICATION:
#   Uses Managed Identity — no API key needed.
#   The Function App's identity must be granted the
#   "Cognitive Services OpenAI User" role on the Azure OpenAI resource.
#
# HOW THE ANALYSIS WORKS:
#   1. We give GPT-4o a "system prompt" (instructions on what to do)
#   2. We give it the transcript text + the SharePoint reference doc text
#   3. GPT-4o returns a JSON object with meeting summary, action items, etc.
#   4. We parse that JSON and return it as a Python dictionary
#   5. That dictionary is used to fill the Word template
#
# USED BY: function_app.py (fn_process_queue)
# =============================================================================

import os      # To read environment variables (AZURE_OPENAI_ENDPOINT, etc.)
import json    # To parse the JSON response from OpenAI
import logging # For log messages
from openai import AzureOpenAI              # The Azure OpenAI Python client library
from azure.identity import DefaultAzureCredential  # For Managed Identity auth


def get_openai_client() -> AzureOpenAI:
    """
    Creates and returns an authenticated Azure OpenAI client.

    AUTHENTICATION METHOD: Managed Identity (no API key)
        We get a token from Azure's identity service and pass it
        to the AzureOpenAI client as "azure_ad_token".
        This token proves we are authorised to use the OpenAI resource.

    REQUIRED APP SETTING:
        AZURE_OPENAI_ENDPOINT → the URL of your Azure OpenAI resource
        Example: "https://my-openai-resource.openai.azure.com/"
        (Found in Azure Portal → your OpenAI resource → Keys and Endpoint)

    RETURNS:
        An AzureOpenAI client object ready to make API calls.

    EXAMPLE — You won't call this directly, it's called by analyse_transcript():
        client = get_openai_client()
        response = client.chat.completions.create(...)
    """

    # Get a Managed Identity token specifically for Azure Cognitive Services
    # (Azure OpenAI is part of Azure Cognitive Services)
    # This is a DIFFERENT scope than Graph API — each Azure service has its own scope
    credential = DefaultAzureCredential()
    token = credential.get_token("https://cognitiveservices.azure.com/.default")

    # Create the OpenAI client with:
    #   azure_endpoint → which OpenAI resource to use (from app settings)
    #   azure_ad_token → our Managed Identity token (instead of an API key)
    #   api_version    → which version of the Azure OpenAI API to use
    return AzureOpenAI(
        azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
        azure_ad_token=token.token,
        api_version="2024-02-01"
    )


def analyse_transcript(transcript_text: str, reference_doc_text: str) -> dict:
    """
    Sends the meeting transcript and reference document to GPT-4o and
    returns a structured analysis as a Python dictionary.

    WHAT THIS FUNCTION DOES:
        1. Connects to Azure OpenAI using get_openai_client()
        2. Sends a "system prompt" telling the AI what to do
        3. Sends the transcript text + reference document as user input
        4. GPT-4o analyses them and returns a JSON object
        5. We parse the JSON and return it as a Python dict

    PARAMETERS:
        transcript_text    → the full VTT transcript text from Teams
                             (fetched by graph_client.get_transcript_content())
                             Example: "WEBVTT\n\n00:00:05.000 --> 00:00:08.000\n<v John>Hello..."

        reference_doc_text → text extracted from the SharePoint template/reference document
                             (used to give OpenAI context about what the meeting was about)
                             Example: "Project Kickoff Document\nObjectives: ..."

    RETURNS a Python dict like:
        {
            "meeting_title": "Weekly Team Standup",
            "meeting_date": "2024-01-15",
            "attendees": ["John Smith", "Jane Doe", "Bob Wilson"],
            "executive_summary": "The team discussed Q1 progress and identified two blockers.",
            "key_decisions": [
                "Delay release to Feb 1st",
                "Hire two more engineers"
            ],
            "action_items": [
                {"owner": "Jane Doe", "task": "Update project plan", "due_date": "2024-01-20"},
                {"owner": "Bob Wilson", "task": "Send status report", "due_date": "TBD"}
            ],
            "discussion_topics": [
                {"topic": "Q1 progress", "summary": "Team is 80% on track."},
                {"topic": "Hiring", "summary": "Approved headcount for 2 engineers."}
            ],
            "reference_doc_relevance": "The reference document outlines project goals...",
            "risks_and_issues": ["Dependency on external vendor", "Budget constraint"],
            "next_steps": ["Schedule follow-up for Feb 1st", "Update Jira tickets"]
        }

    REQUIRED APP SETTING:
        AZURE_OPENAI_DEPLOYMENT → the name of your GPT-4o deployment
        Example: "gpt-4o"
        (Found in Azure OpenAI Studio → Deployments)

    NOTE ON TEXT LIMITS:
        We cap the transcript at 8000 characters and the reference doc at 3000 characters.
        This is because GPT-4o has a limit on how much text it can process at once
        (called "context window"). For very long transcripts, only the first 8000
        characters are sent. You can increase this but it will cost more.
    """

    # Get the authenticated OpenAI client
    client = get_openai_client()

    # Get the deployment name from app settings
    # This is the name you gave your GPT-4o deployment in Azure OpenAI Studio
    deployment = os.environ["AZURE_OPENAI_DEPLOYMENT"]

    # ── System Prompt ──────────────────────────────────────────────────────────
    # The system prompt is like a job description for the AI.
    # It tells GPT-4o EXACTLY what role to play and what format to return.
    # "response_format: json_object" below forces GPT-4o to return valid JSON only.
    system_prompt = """
You are an expert meeting analyst. You will receive:
1. A Teams meeting transcript in VTT format (includes speaker names and timestamps)
2. A reference document from SharePoint (provides context about the meeting topic)

Your job is to analyse the transcript and produce a structured JSON summary
that will be used to automatically fill a Word document report template.

IMPORTANT RULES:
- Return ONLY valid JSON — no markdown code blocks, no explanations, no preamble
- If you cannot find certain information in the transcript, use empty strings or empty arrays
- Keep summaries concise but complete
- Extract actual names from the transcript for attendees and action item owners

Return this EXACT JSON structure (all fields required):
{
  "meeting_title": "The name/subject of the meeting",
  "meeting_date": "YYYY-MM-DD format",
  "attendees": ["Full Name 1", "Full Name 2"],
  "executive_summary": "2-3 sentence overview of the entire meeting",
  "key_decisions": ["Decision 1 made in the meeting", "Decision 2"],
  "action_items": [
    {
      "owner": "Person responsible",
      "task": "What they need to do",
      "due_date": "YYYY-MM-DD or TBD if not mentioned"
    }
  ],
  "discussion_topics": [
    {
      "topic": "Topic name",
      "summary": "Brief summary of what was discussed"
    }
  ],
  "reference_doc_relevance": "How the SharePoint reference document relates to what was discussed",
  "risks_and_issues": ["Risk or issue mentioned 1", "Risk or issue mentioned 2"],
  "next_steps": ["Next step 1", "Next step 2"]
}
"""

    # ── User Message ───────────────────────────────────────────────────────────
    # The user message is the actual data we want GPT-4o to analyse.
    # We include both the transcript and the reference document.
    # [:8000] means "first 8000 characters only" (to stay within token limits)
    user_message = f"""
## MEETING TRANSCRIPT (VTT format):
{transcript_text[:8000]}

## REFERENCE DOCUMENT FROM SHAREPOINT:
{reference_doc_text[:3000]}

Please analyse the transcript above and return the JSON structure.
"""

    # ── Make the OpenAI API Call ───────────────────────────────────────────────
    # This is the actual call to GPT-4o
    response = client.chat.completions.create(
        model=deployment,       # Which GPT-4o deployment to use

        messages=[
            # "system" role = instructions to the AI (how to behave)
            {"role": "system", "content": system_prompt},
            # "user" role = the actual input data (transcript + reference doc)
            {"role": "user", "content": user_message}
        ],

        # temperature=0.1 → makes output more deterministic/consistent
        # (0.0 = always same answer, 1.0 = more creative/random)
        # For data extraction we want consistency, so we use 0.1
        temperature=0.1,

        # max_tokens = maximum length of the response
        # 2000 tokens ≈ ~1500 words — enough for our JSON structure
        max_tokens=2000,

        # response_format="json_object" forces GPT-4o to return ONLY valid JSON
        # Without this, it might add explanations or markdown around the JSON
        response_format={"type": "json_object"}
    )

    # ── Parse the Response ─────────────────────────────────────────────────────
    # response.choices[0].message.content is the raw text response from GPT-4o
    # It will look like: '{"meeting_title": "Standup", "attendees": [...], ...}'
    raw = response.choices[0].message.content
    logging.info(f"OpenAI response received ({len(raw)} characters)")

    # json.loads() converts the JSON string into a Python dictionary
    # If GPT-4o returned invalid JSON (rare with json_object mode), this will raise an exception
    return json.loads(raw)
