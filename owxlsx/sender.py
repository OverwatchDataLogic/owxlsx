import pymongo


def send_data_to_server(host, port, data):
    client = pymongo.MongoClient(host, port)
    db = client['ow']
    ow_data = db['ow_data']
    ow_data.insert_one(data)
    client.close()
