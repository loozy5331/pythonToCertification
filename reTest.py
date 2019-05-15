import os
content = os.popen("git diff HAED HEAD~1").readlines()
print(content)