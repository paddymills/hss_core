import numpy as np

class Plate:

    def __init__(self, ordered_as=None, **kwargs):
        self._init_kwargs = dict()
        for k, v in kwargs.items():
            _key = k.replace(' ', '_').lower()
            setattr(self, _key, v)
            self._init_kwargs[_key] = v

        if ordered_as:
            self = ordered_as

        self._nest = np.full((self.wid, self.len), np.nan, dtype=np.int8)

    def __hash__(self):
        return hash(self.mm)

    def __eq__(self, other):
        return self.mm == other.mm

    def __ne__(self, other):
        return not(self == other)

    def __repr__(self):
        return '{} ({})'.format(self.mm, self.qty)

    def nest_part(self, part_to_nest):
        pass

    def split_plate(self, *split_at):
        plates = list()
        previous_cut = 0
        for cut_at in sorted(split_at):
            self._init_kwargs['len'] = cut_at
            plates.append(Plate(**self._init_kwargs))

        return plates
