def asdfg(n):
    for j in ['a', 'b', 'c', 'd', 'e']:
        for i in range(100):
            if i == n:
                return
            print(f"{j}-{i}")


asdfg(7)

