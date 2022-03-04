def decorator_cache(some_func):
    mem = {}

    def wrapper(*args):
        if args in mem:
            print('returning from cache')
            return mem[args]
        else:
            result = some_func(*args)
            mem[args] = result
            return result

    return wrapper


@decorator_cache
def func(x):
    print("Waiting, i'm calculating")
    return x * x + 1


if __name__ == '__main__':
    print(func(10))
    print(func(10))
