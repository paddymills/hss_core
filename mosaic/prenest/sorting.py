
from itertools import groupby


def create_subsets(parts, plates):
    """
    Pears data down to subsets of connected data.
    (lines of flanges connected by at least one material master)
    """

    lines = create_lines()

    subsets = list()
    while lines:
        new_subset = list()

        new_subset.append(lines.pop())
        while 1:
            subset_mms = list()
            for l in new_subset:
                for mm in l.material_masters:
                    if mm not in subset_mms:
                        subset_mms.append(l)

        for mm in _line.material_masters:
            for i, l in enumerate(lines):
                if l.material_master == mm:
                    new_subset.append(lines.pop(i))


def create_lines(parts):
    """Creates lines of connected flange plates."""
    groups = groupby(parts, lambda part: part.lgf)

    for group in groups:
        yield groups.sort(key=lambda part: part.plate_number)


def line_group_flange(parts):
    return parts.sort(key=lambda part: (part.line, part.group, part.flange))


def line_flange_group(parts):
    return parts.sort(key=lambda part: (part.line, part.flange, part.group))


def radii(parts):
    return parts.sort(key=lambda part: part.radius)


def mirrored(parts):
    pass
