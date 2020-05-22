import re

RANGE_REGEX = re.compile(r"(?P<start>[A-Z0-9]+)[-_](?P<end>[A-Z0-9]+)")
ITEM_REGEX = re.compile(r"(?P<prefix>[A-Z]*)(?P<id>[0-9]+)(?P<suffix>[A-Z]*)")

def clean_and_split_input(_string):
    return _string.replace(' ', '').upper().split(',')


def gen_abc(a, b):
    # Generates a sequential list
    # from an alphabetic character range
    # 
    # - bi-directional (a-c or c-a)
    # - bounds can be different lengths
    # 
    # i.e. Y-AB -> [Y, Z, AA, AB]

    # pad bounds to be same length (right-justified)
    max_length = max(map(len, [a, b]))
    a = a.rjust(max_length)
    b = b.rjust(max_length)

    start, end = sorted([a, b])
    if start == ' ':
        return [end]
    abc = [start]
    
    char_list = [ord(x) for x in start]  # convert to ASCII
    while abc[-1] != end:
        char_list[-1] += 1  # increment ASCII for next character

        # reverse through ASCII-array
        # in case of multiple-character increment
        # (i.e. Z->AA)
        for i in range(len(char_list)-1, -1, -1):
            if char_list[i] == 91:  # Z -> AA
                char_list[i] = 65
                char_list[i - 1] += 1
            elif char_list[i] == 33:  # space -> A
                char_list[i] = 65
        
        # convert ASCII-array to string, strip whitespace
        abc.append(''.join(map(chr, char_list)).strip())

    return abc


def range_to_list(input_str):
    items = list()
    for input_item in clean_and_split_input(input_str):
        range_match = RANGE_REGEX.match(input_item)
        
        if range_match:
            item1 = ITEM_REGEX.match(range_match.group('start'))
            item2 = ITEM_REGEX.match(range_match.group('end'))

            prefix      = item1.group('prefix') or item2.group('prefix') or ''
            start, end  = sorted([int(item1.group('id')), int(item2.group('id'))])
            suffixes    = gen_abc(item1.group('suffix'), item2.group('suffix'))

            for suffix in suffixes:
                for num in range(start, end+1):
                    items.append('{}{}{}'.format(prefix, num, suffix))
        else:
            items.append(input_item)
        

    i = 0
    # handle TB & BT (top and bottom flange)
    while i < len(items):
        val = items[i]

        if val.startswith('TB') or val.startswith('BT'):
            prefix = val[:2]
            items.append(val.replace(prefix, 'T', 1))
            items.append(val.replace(prefix, 'B', 1))
            items.remove(val)
        else:
            i += 1

    return items
