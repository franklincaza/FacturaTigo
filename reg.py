import re

txt = "The rain in Spain"
x = re.findall(" \w[in{2}]", txt)
print(x)
