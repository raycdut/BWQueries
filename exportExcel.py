from Mongo_Client import Mongo_Client
import pandas as pd

class ExportExcel():
    def __init__(self):
        self.client = Mongo_Client()
        
    def export(self):
        collectionlist = ['tmallProduct_Tags_Final','Products', 'category_statistic','productDetails']
        for collection_name in collectionlist:
            collection = self.client.db[collection_name]
            data = pd.DataFrame(list(collection.find()))
            del data['_id']

            if collection_name == 'productDetails':
                print('process args')
                data['args'] = data['args'].apply(lambda x: ';'.join(x) )
                data['spuId'] = data['spuId'].apply(lambda x: ';'.join(x))
                data['sellerId'] = data['sellerId'].apply(lambda x: ';'.join(x))
                data = data.drop_duplicates(subset='productId',keep='first')

            
            
            data.to_excel(collection_name+'.xlsx', sheet_name='sheet1', index=False)

    def exportOpsLinks(self):
        collection = self.client.db['allinone']
        data = pd.DataFrame(list(collection.find()))
        del data['_id']
        data.to_excel('allinone.xlsx', sheet_name='sheet1', index=False)

    def export_comments(self):
        products = self.client.db['Products']
        cateogries = products.distinct('category')
        for category in cateogries:
            print(category)
            data = pd.DataFrame(list(self.client.db[category].find()))
            del data['_id']
            data = data.drop_duplicates(keep='first')
            data.to_excel(category+'.xlsx', sheet_name='sheet1', index=False)
if __name__ == "__main__":
    ee = ExportExcel()
    #ee.export()
    ee.exportOpsLinks()
    #ee.export_comments()
