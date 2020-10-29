import multiprocessing as mp
import os
import psutil


def iterate(iterator) -> None:
    while True:
        try:
            next(iterator)
        except StopIteration:
            break


def do_nothing():
    pass


def _limit_cpu():
    p = psutil.Process(os.getpid())
    p.nice(psutil.IDLE_PRIORITY_CLASS)


def multiprocess(func, lst=None, *args):
    def log_result(ret_val):
        lock.acquire()
        results.append(ret_val)
        lock.release()

    print(func)
    print(lst)

    lock = mp.Lock()
    pool = mp.Pool(None, _limit_cpu)
    results = list()

    for item in lst:
        if args:
            arg_s = (item, args)
        else:
            arg_s = item
        pool.apply_async(func, args=arg_s, callback=log_result)

    pool.close()
    pool.join()

    if results:
        return results
    else:
        raise RuntimeError
