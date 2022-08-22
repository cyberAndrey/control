from app import db


class Key(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    kvantum = db.Column(db.String(64), index=True, unique=True)
    taken = db.Column(db.Boolean, default=False, nullable=False)
    who = db.Column(db.String(64), index=True, default=None)

    def __repr__(self):
        return '{}'.format(self.kvantum)


class Staff(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), index=True)

    def __repr__(self):
        return '{}'.format(self.name)


class Guest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), index=True)

    def __repr__(self):
        return '{}'.format(self.name)
