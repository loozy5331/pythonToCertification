import os
content = os.popen("git diff HAED~1 HEAD").readlines()
print(content)