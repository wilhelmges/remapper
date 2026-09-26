from waste_keeper.domain.core import remove_duplicate_spaces
from waste_keeper.waste_status import Waste_status

class StatusDetector:
    def __init__(self, descr):
        self.descr = descr
        if isinstance(descr, str):
            self.descr = (remove_duplicate_spaces(descr.strip().lower()).
                          replace("зміни","змін"))

    @property
    def waste_status(self):
        if self.descr is None:
            return Waste_status.BLANK
        if not isinstance(self.descr, str):
            return Waste_status.ETC
        if self.descr.startswith("єас") or "інспекторське" in self.descr:
            return Waste_status.WROTEOFF
        elif (self.descr.startswith("змін") or  self.descr.startswith("змінено")
              or "змін наказ" in self.descr or "внесення змін" in self.descr):
            return Waste_status.UPDATED
        elif (self.descr.startswith("скасування") or self.descr.startswith("скасовано")or "втрату" in self.descr
              or "втратив" in self.descr or "втрата чинності" in self.descr):
            return Waste_status.CANCELED
        elif "розслідування" in self.descr or "рослідування" in self.descr or "розслідуваня" in self.descr:
            return Waste_status.INQUIRY
        elif ("винн" in self.descr) or ("стягнути" in self.descr):
            return Waste_status.GUILTED
        elif len(self.descr) > 1:
            return Waste_status.ETC
        else:
            return Waste_status.BLANK


if __name__=='__main__':
    sd = StatusDetector('вважати таким що втратив чинність на підставі наказу №907 від 10.02.2026')
    print(sd.waste_status, sd.waste_status.value)

#стягнути з Бухнацевич Д.В.
