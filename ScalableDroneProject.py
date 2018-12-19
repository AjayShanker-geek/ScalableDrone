import math, sys, pip

class Motor:
    
    def __init__(self, modelName, mass, baseDiameter, holeSize, holeDistanceXaxis, holeDistanceYaxis):
        self.modelName = modelName
        self.mass = mass
        self.measurments = self.Measurments( baseDiameter, holeSize, holeDistanceXaxis, holeDistanceYaxis)
        

    def show(self):
        print("\n\n\n\nMotor Model: {0}\nMass: {1}g".format( self.modelName, self.mass))
        self.measurments.show()
        

        
    class Measurments:
        
        def __init__(self, baseDiameter, holeSize, holeDistanceXaxis, holeDistanceYaxis):
            self.baseDiameter = baseDiameter
            self.holeSize = holeSize
            self.holeDistanceXaxis = holeDistanceXaxis
            self.holeDistanceYaxis = holeDistanceYaxis

        def show(self):
            print("Base Diameter: {0}mm\nHole Size: {1}mm\nHole Distance X axis: {2}mm\nHole Distance Y axis: {3}mm\n".format( self.baseDiameter, self.holeSize, self.holeDistanceXaxis, self.holeDistanceYaxis))

class Camera:

    def __init__(self, modelName, mass, frameRate, resolution, length, width, height):
        self.modelName = modelName
        self.mass = mass
        self.frameRate = frameRate
        self.resolution = resolution
        self.measurments = self.Measurments( length, width, height)

    def show(self):
        print("\n\n\n\nCamera Model: {0}\nMass: {1}g\nFrame Rate: {2}fps\nResolution: {3}p".format( self.modelName, self.mass, self.frameRate, self.resolution))
        self.measurments.show()


    class Measurments:

        def __init__( self, length, width, height):
            self.length = length
            self.width = width
            self.height = height

        def show(self):
            print("Length: {0}mm\nWidth: {1}mm\nHeight: {2}mm\n".format( self.length, self.width, self.height))

class FlightController:

    def __init__(self, modelName, mass, length, width, height):
        self.modelName = modelName
        self.mass = mass
        self.measurments = self.Measurments( length, width, height)

    def show(self):
        print("FCC Model: {0}\nMass: {1}g".format( self.modelName, self.mass))
        self.measurments.show()

    class Measurments:

        def __init__(self, length, width, height):
            self.length = length
            self.width = width
            self.height = height

        def show(self):
            print("Length: {0}mm\nWidth: {1}mm\nHeight: {2}mm".format( self.length, self.width, self.height))
                
class ElectronicSpeedController:

    def __init__(self, modalName, mass, )        
        




#Motor Properties            
motor = Motor('1105-7500KV', 5.9, 27.5, 3, 16, 19)

#Camera Properties
camera = Camera('GWY CMOS Camera with VCR', 20, 60, 720, 26, 28, 25)

#Flight Controller Properties
flightCC = FlightController('PixRacer R15', 10.54, 36, 36 , 11.5)

motor.show()
camera.show()
flightCC.show()
            
            
