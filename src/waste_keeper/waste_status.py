from enum import Enum

class Waste_status(Enum):
    BLANK = 'blank'
    CANCELED = "canceled"
    UPDATED = "updated"
    WROTEOFF = "wroteoff"
    GUILTED = "guilted"
    INQUIRY = "inquiry"
    ETC = "etc"