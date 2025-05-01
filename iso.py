import pandas as pd

data = pd.read_excel('Missing_Buses.xlsx')

# Example: your list of SPP values
spp_list = {
    "MJMEUC", "AECC", "SWPA", "AEPW", "GRDA", "OKGE", "WFEC", "SPS", "OMPA", "MIDW", 
    "SUNC", "MKEC", "WERE", "GMO", "KCPL", "KACY", "EMDE", "INDN", "SPRM", "NPPD",
    "MEAN", "GRIS", "OPPD", "LES", "WAPA", "BEPC", "HCPD", "CBPC", "NWPS", "MRES",
    "BEPC-SPP", "WBDC-WE", "ERCOT", "WECC"
}

pjm_list = {
    "AP", "ATSI", "AEP", "OVEC", "DAY", "DEO&K", "DLCO", "CE", "PJM",
    "PENELEC", "METED", "JCP&L", "PPL", "PECO", "PSE&G", "BGE", "PEPCO",
    "AE", "DP&L", "UGI", "RECO", "SMECO", "EKPC", "DVP", "APS", "PSEG", "ME"
}

miso_list = {
    "HE", "DEI", "SIGE", "IPL", "NIPS", "METC", "ITCT", "WEC", "MIUP", "CLOV",
    "BREC", "HMPL", "EES-EMI", "EES-EAI", "LAGN", "CWLD", "SMEPA", "EES",
    "AMMO", "AMIL", "CWLP", "SIPC", "GLH", "CLEC", "LAFA", "LEPA", "XEL", "MUNI",
    "MMPA", "CMMPA", "MP", "SMPPA", "GRE", "OTP", "MPC", "MRES", "ALTW", "MPW",
    "MEC", "RPGI", "IAMU", "MMEC", "MDU", "BEPC-SPP", "BEPC-MISO", "DPC", "WPPI",
    "ALTE", "WPS", "CWP", "MEWD", "MPU", "MGE", "UPPC"
}

serc_list = {
    "AECI", "CPLE", "CPLW", "DUK", "SCEG", "SCPSA", "SOUTHERN", "TVA", "YAD",
    "SEHA", "SERU", "SETH", "LGEE", "OMUA", "SMT", "TAP"
}

frcc_list = {
    "FPLNW", "FPL", "PEF", "FTP", "GVL", "HST", "JEA", "KEY", "LWU", "NSB", "FMPP",
    "SEC", "TAL", "TECO", "FMP", "NUG", "RCU", "TCEC", "OSC", "OLEANDER", "CALPINE",
    "HPS", "DESOTOGEN", "IPP-REL"
}

nyiso_list = {
    "WEST",
    "GENESEE",
    "CENTRAL",
    "NORTH",
    "MOHAWK",
    "CAPITAL",
    "HUDSON",
    "MILLWOOD",
    "DUNWOODIE",
    "NYC",
    "L ISLAND"
}

# Now apply a function to the column
def map_region(value):
    if value in spp_list:
        return "SPP"
    elif value in pjm_list:
        return "PJM"
    elif value in miso_list:
        return "MISO"
    elif value in serc_list:
        return "SERC"
    elif value in frcc_list:
        return "FRCC"
    elif value in nyiso_list:
        return "NYISO"
    else:
        return value
    
data[' Area Name'] = data[' Area Name'].apply(map_region)

data.to_excel('Missing_Buses.xlsx')