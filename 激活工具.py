# activate.py（仅发给付费用户）
from FK import activate

if activate():
    print("激活成功！程序已解锁")
else:
    print("激活失败，请确保程序已正常运行过一次")