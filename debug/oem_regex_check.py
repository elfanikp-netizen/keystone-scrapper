import re

sample = "OEM 1234567\nOEM A12345\nOEM 12345A\nInterchange 123-45678"


def old_logic(text):
    tokens = []
    for token in re.findall(r"[A-Za-z0-9-]+", text):
        token = token.strip()
        if re.search(r"(?i)(?:[A-Za-z]+\d{5,7}|\d{3}\d{5,7})", token):
            tokens.append(token)
    return tokens

result = old_logic(sample)
print(result)
assert "1234567" in result or "12345A" in result or "A12345" in result
