import json


class BaseModel(object):
    def to_json(self):
        dump = json.dumps(self, default=lambda o: o.__dict__, sort_keys=True)
        return json.loads(dump)