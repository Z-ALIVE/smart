'''
    Class eSight: contains all eSight IPs
    Class FM: contains all FusionSphere Management IP
'''
class eSight():

    def __init__(self):
        self.nl1 = '10.90.194.74'
        self.nl2 = '10.90.92.16'
        self.sl1 = '10.90.210.10'
        self.sl2 = '10.90.158.10'
        self.nm1 = '10.90.118.16'
        self.nm2 = '10.90.134.10'
        self.sm1 = '10.90.20.170'
        self.sm2 = '10.90.177.16'
        self.vis1 = '10.90.52.170'
        self.vis2 = '10.85.78.10'
        self.min1 = '10.85.74.10'
        self.min2 = '10.90.74.16'

class FM():

    def __init__(self):
        self.nl1 = '10.90.194.68'
        self.nl2 = '10.90.92.4'
        self.sl1 = '10.90.210.4'
        self.sl2 = '10.90.158.4'
        self.nm1 = '10.90.134.4'
        self.nm2 = '10.90.118.4'
        self.sm1 = '10.90.20.164'
        self.sm2 = '10.90.177.4'
        self.vis1 = '10.90.52.164'
        self.vis2 = '10.85.78.4'
        self.min1 = '10.85.74.4'
        self.min2 = '10.90.74.4'

esight_ip = eSight()
nl1 = esight_ip.nl1
nl2 = esight_ip.nl2
sl1 = esight_ip.sl1
sl2 = esight_ip.sl2
nm1 = esight_ip.nm1
nm2 = esight_ip.nm2
sm1 = esight_ip.sm1
sm2 = esight_ip.sm2
vis1 = esight_ip.vis1
vis2 = esight_ip.vis2
min1 = esight_ip.min1
min2 = esight_ip.min2
sites = [nl1, nl2, sl1, sl2, nm1, nm2, sm1, sm2, vis1, vis2, min1, min2]
sites_n = ['nl1', 'nl2', 'sl1', 'sl2', 'nm1', 'nm2', 'sm1', 'sm2', 'vis1', 'vis1', 'vis2', 'min1', 'min2']
# sites = [nl1, sm2]
# sites_n = ['nl1', 'sm2']