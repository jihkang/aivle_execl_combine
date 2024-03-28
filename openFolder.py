import os

# with os.scandir(os.path(__dir__)) as it:
# 	for entry in it:
# 		print(entry.name)


def getPath(func):
    def wrapper(*args, **kwargs):
        if len(args) > 0:
            path = args[0]
        else:
            path = os.getcwd() + '/excel'
        return func(path)
    return wrapper


@getPath
def scanDir(path):
    with os.scandir(path) as it:
        ret = []
        for entry in it:
            if ('.xlsx' in entry.name):
                ret.append(path + '/' + entry.name)
        return ret


if __name__ == "__main__":
    data = scanDir()
    print(data)
