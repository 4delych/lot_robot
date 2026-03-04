import sys
import time

from openai import OpenAI


API_KEY = "OWMyYjA4ZjItNGY2Ni00OTNjLWJlMmUtN2Y5YTI1MjYwYWNi.9b5fe2f4b7aa0c1358219ee59f9b2b25"
BASE_URL = "https://foundation-models.api.cloud.ru/v1"
MODEL = "Qwen/Qwen3-235B-A22B-Instruct-2507"
PROVIDER = "cloudru"


def main() -> int:
    prompt = "Ответь одним словом: работает"
    if len(sys.argv) > 1:
        prompt = " ".join(sys.argv[1:]).strip() or prompt

    client = OpenAI(
        api_key=API_KEY,
        base_url=BASE_URL,
    )

    print(f"provider={PROVIDER}")
    print(f"base_url={BASE_URL}")
    print(f"model={MODEL}")
    print(f"prompt={prompt}")

    started_at = time.time()
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "user", "content": prompt},
            ],
            max_tokens=64,
            temperature=0.0,
            timeout=60,
        )
    except Exception as e:
        elapsed = time.time() - started_at
        print(f"request_failed_after={elapsed:.2f}s")
        print(f"error={type(e).__name__}: {e}")
        return 2

    elapsed = time.time() - started_at
    content = ""
    try:
        content = (response.choices[0].message.content or "").strip()
    except Exception:
        content = ""

    print(f"request_ok_after={elapsed:.2f}s")
    print(f"response_id={getattr(response, 'id', '')}")
    print(f"content={content}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
