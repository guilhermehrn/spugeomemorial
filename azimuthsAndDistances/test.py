

pol = [[[0.0, 0.0], [4.0, 0.0], [4.0, 4.0], [0.0, 4.0], [0.0, 0.0]]]


def genericCalculAzimuthDistance(self):

    geom = self.geom
    dim = len(geom)

    for i in range(0, dim):
        polygon  = geom[i]
        dimPolygon = len(geom[i])
        for ponto1 in range(0, dimPolygon):
            pontIndex = polygon.index(ponto1)
            ponto2 = polygon[pontoIndex+1]

            print("Par", ponto1, ponto2)



