import contextlib
import time
import datetime


@contextlib.contextmanager
def time_elapsed(loop, obj):
    cur = loop.time()
    yield
    obj.elapsed = loop.time() - cur
    #print('time_elapsed : {}, {}'.format(cur, obj.elapsed))

def now():
    #m = datetime.datetime.now().strftime('%M')
    m = datetime.datetime.now()
    return m
    
