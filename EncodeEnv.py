import base64

with open(".env", "rb") as f:
    encoded_env = base64.b64encode(f.read()).decode("utf-8")

print(f"Encoded .env:\n{encoded_env}")
