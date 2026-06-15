from core import calculate_md5
from pathlib import Path
from difflib import SequenceMatcher
from get_sigmatures_from_xlsx import get_list_of_sigmatures_from_xlsx

def compare_signatures(old_rows, new_rows):

    deleted = []
    modified = []
    added = []

    matcher = SequenceMatcher(None, old_rows, new_rows)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():

        if tag == "equal":
            continue

        elif tag == "delete":

            # видалені рядки з першого списку
            deleted.extend(range(i1, i2))

        elif tag == "insert":

            # додані рядки у другому списку
            added.extend(range(j1, j2))

        elif tag == "replace":

            old_count = i2 - i1
            new_count = j2 - j1

            common = min(old_count, new_count)

            # рядки, які змінилися
            modified.extend(range(i1, i1 + common))

            # якщо старих більше -> частина була видалена
            if old_count > new_count:
                deleted.extend(range(i1 + common, i2))

            # якщо нових більше -> частина була додана
            elif new_count > old_count:
                added.extend(range(j1 + common, j2))

    return {
        "deleted": deleted,
        "modified": modified,
        "added": added,
    }

if __name__=='__main__':
    oldfile = Path(r"./backups/накази_втрати майна  А4007.xlsx")
    newfile = Path(r"./backups/накази_втрати майна  А4007_c09d904c.xlsx")
    if calculate_md5(oldfile) == calculate_md5(newfile):
        print('equel files')
        exit()

    old = get_list_of_sigmatures_from_xlsx(oldfile)
    new = get_list_of_sigmatures_from_xlsx(newfile)

    result = compare_signatures(old, new)

    print(result)

    #print(calculate_md5())
    #print(calculate_md5())