from peewee import *


db = SqliteDatabase('acounting-db.db')


class CheckModel(Model):
    obj_id = CharField(unique=True, index=True)
    number = SmallIntegerField()
    amount = IntegerField()
    recieved_docs = CharField()
    condition = CharField()
    date_check = DateField()
    date_recieved_ckeck = DateField()
    bank_name = CharField(null=True)
    submit_date = DateField(null=True)

    class Meta:
        database = db # This model uses the "people.db" database.


def initial_db():
    db.connect()
    db.create_tables([CheckModel])
    db.close()
