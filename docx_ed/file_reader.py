import asyncio
import os
import json


class FileReader:
    def __init__(self, name: str, ):
        self.name = name

    async def read_file_from_pre(self):
        with open(f"../files/gost/{self.name}", encoding='UTF-8') as f:
            return json.load(f)

    async def read_file_from_user(self):
        with open(f"../files/user_json/{self.name}", encoding='UTF-8') as f:
            return json.load(f)

    @classmethod
    def get_actual_pre_gosts(cls):
        return {gost.replace('.json', ''): gost for gost in os.listdir('../files/gost/')}

    @classmethod
    def get_user_gosts(cls):
        return {gost.replace('.json', ''): gost for gost in os.listdir('../files/user_json/')}


if __name__ == '__main__':
    js_dict = FileReader('user.json')
    file_cont = asyncio.run(js_dict.read_file())
    print(file_cont)
    print(js_dict.get_actual_pre_gosts())
