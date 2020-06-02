
from string import Template


class CountingIter:

    def __init__(self, iter, caption="", total=False):
        self._iter = iter
        self._index = 0
        self.template = Template("\r[$index] {}".format(caption))

        if total:
            self.template = Template(
                "\r[$index/{}] {}".format(len(iter), caption))

    def __iter__(self):
        return self

    def __next__(self):
        if self._index == len(self._iter):
            print()
            raise StopIteration

        elem = self._iter[self._index]
        self._index += 1

        print(self.template.substitute(index=self._index), end='')

        return elem
