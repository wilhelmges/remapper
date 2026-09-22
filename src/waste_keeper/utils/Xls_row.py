import inspect

class XlsRow:
    @property
    def status_description(self):
        raise Exception("Abstract method. Not implemented")

    def cell(self, col: int):
        return self.ws.cell(row=self.row, column=col)

    def strike(self, col):
        if self.ws.cell(row=self.row, column=col).font.strike:
            return True
        return False

    def __contains__(self, items: list|str):
        d = self.status_description
        if d is None or (not isinstance(d, str)):
            return False
        d = (str(d)).lower()
        if isinstance(items, (list, tuple, set)):
            return any(item.lower() in d for item in items)
        item = items
        return item in d

    def __repr__(self):
        props = []
        for name, _ in inspect.getmembers(type(self), lambda x: isinstance(x, property)):
            try:
                value = getattr(self, name)
            except Exception as e:
                value = f"<Error: {e}>"
            props.append(f"{name}={value!r}")
        return f"{type(self).__name__}({', '.join(props)})"