'''
    Class eSight: contains all eSight IPs
    Class FM: contains all FusionSphere Management IP
'''
class eSight():

    def __init__(self):
        self.nl1 = '*'
        self.nl2 = '*'
        self.sl1 = '*'
        self.sl2 = '*'
        self.nm1 = '*'
        self.nm2 = '*'
        self.sm1 = '*'
        self.sm2 = '*'
        self.vis1 = '*'
        self.vis2 = '*'
        self.min1 = '*'
        self.min2 = '*'

class FM():

    def __init__(self):
        self.nl1 = '*'
        self.nl2 = '*'
        self.sl1 = '*'
        self.sl2 = ''
        self.nm1 = ''
        self.nm2 = ''
        self.sm1 = ''
        self.sm2 = ''
        self.vis1 = ''
        self.vis2 = ''
        self.min1 = ''
        self.min2 = ''

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
