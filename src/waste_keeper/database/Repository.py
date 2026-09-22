class Repository:

    def __init__(self, model, session):
        self.model = model
        self.session = session

    def create(self, obj):
        self.session.add(obj)
        self.session.commit()
        self.session.refresh(obj)
        return obj

    def get(self, id):
        return self.session.get(self.model, id)