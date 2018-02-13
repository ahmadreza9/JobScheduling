

class test:
    def __init__(self):
        self.num = 1
    def s(self,a):
        self.num = a
    def g(self):
        return self.num
def t(a):
    a.s(2)
t2 = test()
t(t2)
print(t2.g())

for i in reversed(list(range(9))):
    print(i)