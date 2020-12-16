from progress.bar import Bar
import time
print("starting...")
bar= Bar("邮件发送中...",max=100)
for i in range(100):
    time.sleep(0.1)
    bar.next()
bar.finish()
print('finished')