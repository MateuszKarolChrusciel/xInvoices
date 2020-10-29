import datetime
import time

from common.program_objs import Time

def timeit(method):
    def timed(*args, **kw):
        ts = time.time()
        result = method(*args, **kw)
        te = time.time()

        t = (te - ts)

        h = int(t // 3600)
        m = int(t // 60)
        s = int(t // 1)
        ms = int(t % 1000)

        rt = Time(h, m, s, ms)

        print(f"\nMethod \'{method.__name__}\' executed in {str(rt)}")

        return result

    return timed
