def remove_parentheses(text):
    result = []
    depth = 0

    for char in text:
        if char == '(':
            depth += 1
        elif char == ')':
            if depth > 0:
                depth -= 1
        elif depth == 0:
            result.append(char)

    return ' '.join(''.join(result).split())

def get_parentheses_text(s: str) -> str | None:
    start = s.find("(")
    if start == -1:
        return None

    level = 0

    for i in range(start, len(s)):
        if s[i] == "(":
            level += 1
        elif s[i] == ")":
            level -= 1

            if level == 0:
                return s[start:i + 1]

    return None

