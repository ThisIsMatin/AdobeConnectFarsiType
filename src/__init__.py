import sys
try:
    from FarsiType import AdobeConnectFarsiType
except ImportError as e:
    input('Cannot import modules, please insall requirements.txt')
    sys.exit(0)
appliaction = AdobeConnectFarsiType()