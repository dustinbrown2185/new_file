# Regular expression

import re

description = "The Norwegian Blue is a wonderful parrot. This parrot is notable for its exquisite plumage."
pattern = "(parrot)"

replacement = "ex-\\1"
print(re.sub(pattern, replacement, description))

