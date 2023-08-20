import json
import requests

class eSight():
    '''
            param host: eSight IP management
            param dn: device unique identifier
            param servertype: server type (rack/blade)
    '''

    def __init__(self, host):

        self.host = host
        self.username = 'openApiUser'
        self.password = 'Changeme_123'

    def get_openid(self):
        requests.packages.urllib3.disable_warnings()
        url = 'https://{}:32102/rest/openapi/sm/session'.format(self.host)
        data = {
            'userid': self.username,
            'value': self.password,
        }
        response = requests.put(url=url, json=data, verify=False)
        result = response.text
        content = json.loads(result)
        openid = content['data']
        return openid


    def get_serverlist(self, openid, servertype):
        url = 'https://{}:32102/rest/openapi/server/device'.format(self.host)
        header = {
            'openid' : openid
        }
        payload = {
            'servertype' : servertype
        }
        try:
            response = requests.request("GET", url=url, headers=header, json=payload, verify=False)
            response = json.dumps(response.json(), sort_keys=True, indent=4)
            return response
        except:
            return None

    def get_serverdetails(self, openid, dn):
        url = 'https://{}:32102/rest/openapi/server/device/detail'.format(self.host)
        header = {
            'openid' : openid
        }
        payload = {
            'dn' : dn
        }
        try:
            response = requests.request("GET", url=url, headers=header, json=payload, verify=False)
            response = json.dumps(response.json(), sort_keys=True, indent=4)
            return response
        except:
            return None

    def get_networklist(self, openid):
        url = 'https://{}:32102/rest/openapi/network/nedevice'.format(self.host)
        header = {
            'openid' : openid
        }
        payload = {}
        try:
            response = requests.request("GET", url=url, headers=header, json=payload, verify=False)
            response = json.dumps(response.json(), sort_keys=True, indent=4)
            return response
        except:
            return None

    def get_storagelist(self, openid):
        url = 'https://{}:32102/rest/openapi/storage/device'.format(self.host)
        header = {
            'openid' : openid
        }
        payload = {
            'deviceSeries' : 'unifiedstor'
        }
        try:
            response = requests.request("GET", url=url, headers=header, json=payload, verify=False)
            response = json.dumps(response.json(), sort_keys=True, indent=4)
            return response
        except:
            return None

    def get_storagedisklist(self, openid, dn):
        url = 'https://{}:32102/rest/openapi/storage/physicaldisk'.format(self.host)
        header = {
            'openid' : openid
        }
        payload = {
            'dn' : dn
        }
        try:
            response = requests.request("GET", url=url, headers=header, json=payload, verify=False)
            response = json.dumps(response.json(), sort_keys=True, indent=4)
            return response
        except:
            return None



if __name__ == '__main__':
    eSight.get_openid()
    eSight.get_serverlist()
    eSight.get_serverdetails()
    eSight.get_openid()
    eSight.get_networklist()
    eSight.get_storagelist()

