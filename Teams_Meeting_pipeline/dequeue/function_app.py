"""
================================================================================
 ENTRY POINT  ->  function_app.py
================================================================================

WHAT IS THIS FILE?
------------------
In the Azure Functions "Python v2 programming model", THIS single file is the
entry point that the Azure runtime loads. You declare your functions here using
Python decorators (the lines that start with `@app.`). There is no separate
`function.json` per function like in the older v1 model - the decorators
generate that configuration for you.

WHAT IS A "FUNCTION"?
---------------------
An Azure Function is just a Python function that the cloud runs FOR you when
something happens. That "something" is called a TRIGGER. Examples of triggers:
  - an HTTP request arrives        (HTTP trigger)
  - a message lands on a queue     (Queue trigger)
  - a timer ticks                  (Timer trigger)
You do not run the function yourself; Azure runs it when the trigger fires.

WHAT THIS FILE DEFINES
----------------------
Two functions that BOTH do the exact same work by calling one shared pipeline:
  1. manual_trigger  - an HTTP trigger you call yourself during development.
  2. queue_trigger   - a Queue trigger Azure calls automatically in production.

Keeping the real logic in `shared/processor.py` means we never copy-paste it.
================================================================================
"""

import json
import logging

# The Azure Functions SDK. `func` gives us the decorators (func.FunctionApp,
# func.route, func.queue_trigger) and the request/response types.
import azure.functions as func

from shared.config import Config
from shared.processor import process_message

# `FunctionApp` is the object that holds all your functions. The runtime looks
# for a module-level variable named `app` of this type. Every @app.* decorator
# below registers one function with it.
app = func.FunctionApp()


# --------------------------------------------------------------------------- #
# 1. MANUAL (HTTP) TRIGGER - for development & testing
# --------------------------------------------------------------------------- #
# @app.function_name : the name Azure shows for this function in the portal/logs.
# @app.route         : turns this into an HTTP trigger.
#     route="process-transcript" -> URL path .../api/process-transcript
#     methods=["POST"]           -> only POST requests are accepted
#     auth_level=FUNCTION        -> caller must supply a function key (a secret).
#                                   Use ANONYMOUS only if you want it open to all.
@app.function_name(name="manual_trigger")
@app.route(route="process-transcript", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def manual_trigger(req: func.HttpRequest) -> func.HttpResponse:
    """HTTP-triggered function: lets YOU run the pipeline on demand.

    Why it exists:
        While developing you do not want to wait for a real Teams transcript to
        land on the queue. With this you can POST a sample message yourself
        (from VS Code, curl, or Postman) and watch the whole pipeline run.

    Args:
        req: the incoming HTTP request. Azure builds this object for you and
             passes it in. `req.get_body()` gives the raw bytes of the body.

    Returns:
        An HttpResponse. Whatever you return here becomes the HTTP reply the
        caller sees (status code + body).
    """
    # `logging` writes to the Functions log stream / Application Insights.
    # Always log at the start so you can confirm the function actually fired.
    logging.info("manual_trigger invoked.")

    try:
        # The body arrives as bytes; decode to a normal string.
        body = req.get_body().decode("utf-8")
        if not body:
            # 400 = "Bad Request": the caller sent nothing to process.
            return func.HttpResponse("Request body is empty.", status_code=400)

        # We accept two convenient shapes so testing is easy:
        #   (a) the raw queue message JSON directly, OR
        #   (b) {"message": "<stringified json>"} - handy when copying the exact
        #       string that would sit on the queue.
        try:
            parsed = json.loads(body)
            if isinstance(parsed, dict) and "message" in parsed:
                message = parsed["message"]
                if not isinstance(message, str):
                    message = json.dumps(message)
            else:
                message = body
        except json.JSONDecodeError:
            # Body was not JSON at all - treat it as the raw message string.
            message = body

        # Hand off to the ONE shared pipeline. This is the important line.
        result = process_message(message)

        # 200 = success. We return a small JSON summary so you can see what
        # happened right in the HTTP response.
        return func.HttpResponse(
            json.dumps({"status": "success", "result": result}),
            status_code=200,
            mimetype="application/json",
        )

    except Exception as exc:  # noqa: BLE001
        # Catch-all so the caller gets a clean error instead of a raw stack
        # trace. `logging.exception` records the full traceback for debugging.
        logging.exception("manual_trigger failed.")
        return func.HttpResponse(
            json.dumps({"status": "error", "error": str(exc)}),
            status_code=500,  # 500 = "Internal Server Error"
            mimetype="application/json",
        )


# --------------------------------------------------------------------------- #
# 2. QUEUE TRIGGER - for production (Azure calls it automatically)
# --------------------------------------------------------------------------- #
# @app.queue_trigger wires this function to an Azure Storage Queue:
#     arg_name="msg"      -> the message is passed in as the `msg` parameter.
#     queue_name=...      -> WHICH queue to watch (we read it from settings).
#     connection="QueueConnection"
#                         -> the NAME of an app setting that holds the storage
#                            account connection string. Azure reads the actual
#                            secret from that setting at runtime - you never put
#                            the secret here in code.
@app.function_name(name="queue_trigger")
@app.queue_trigger(
    arg_name="msg",
    queue_name=Config.QUEUE_NAME,
    connection="QueueConnection",
)
def queue_trigger(msg: func.QueueMessage) -> None:
    """Queue-triggered function: runs once per message, automatically.

    How it behaves:
        Azure delivers each new queue message to this function. We decode it and
        pass it to the same `process_message` pipeline used by the HTTP trigger.

    Important - error handling for queues:
        Notice we DON'T wrap this in try/except. If the pipeline raises, the
        Functions runtime will automatically retry the message a few times, and
        if it keeps failing, move it to a "poison queue" (named
        <queue>-poison). That preserves the failing message so you can inspect
        it later - exactly what you want for auditing. Swallowing the error here
        would silently lose messages.
    """
    # QueueMessage.get_body() returns bytes; decode to text. `msg.id` is a
    # unique id Azure assigns - useful to correlate logs for one message.
    message = msg.get_body().decode("utf-8")
    logging.info("queue_trigger invoked. message_id=%s", msg.id)
    process_message(message)
