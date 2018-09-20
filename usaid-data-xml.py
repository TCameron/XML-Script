import datetime
import time
import sys
import shutil
import os
from xml.etree import ElementTree
from xml.etree.ElementTree import Element, SubElement
from xml.dom import minidom
import pandas

__author__ = "Timothy Cameron"
__email__ = "tcameron@devtechsys.com"
__date__ = "09-20-2018"
__version__ = "0.38"
date = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3]+'Z'


def prettify(elem):
    """
    Return a pretty-printed string for the xml elements.
    :param elem: The current XML tree.
    :return reparsed: The pretty-printed string from the XML tree.
    """
    rough_string = ElementTree.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")


def id_loop(ombfile):
    """
    Return lists of main activities and the ids for each row.
    :param ombfile: The source file to grab the data from
    :return ids: The main activity identifiers
    :return idswawards: The sub-activity identifiers
    :return isos: The ISO-3166 country codes for each activity
    """
    ids = list()
    idswawards = list()
    isos = list()
    # USAID is identified as US-GOV-1 within IATI.
    code = '1'
    for i in range(0, len(ombfile.index)):
        try:
            country = str(int(ombfile["DAC Regional Code"][i]))
        except ValueError:
            if str(ombfile["ISO Alpha Code"][i]) != 'nan':
                try:
                    country = str(ombfile["ISO Alpha Code"][i])
                except ValueError:
                    country = '998'
            else:
                # This is due to Namibia's code being "NA", but numpy counts "NA" as not applicable, so it gets skipped
                try:
                    namibia = str(int(ombfile["DAC Country Code"][i]))
                    if namibia == '275':  # Being used as a temporary stopgap
                        country = 'NA'
                    else:
                        country = '998'
                except ValueError:
                    country = '998'
        awardid = str(omb["Implementing Mechanism ID"][i])
        entry = 'US-GOV' + '-' + code + '-' + country
        ids.append(entry)
        entry += '-' + awardid
        idswawards.append(entry)
        isos.append(country)
    return ids, idswawards, isos


def group_split(ombfile):
    """
    Return lists of separate recipients and the ids for each row.
    :param ombfile: The source file to grab the data from
    :return ids: The main activity identifiers
    """
    actives = list()
    ids = list()
    for i in range(0, len(ombfile.index)):
        try:
            country = str(int(ombfile["DAC Regional Code"][i]))
        except ValueError:
            if str(ombfile["ISO Alpha Code"][i]) != 'nan':
                try:
                    country = str(ombfile["ISO Alpha Code"][i])
                except ValueError:
                    country = '998'
            else:
                try:
                    namibia = str(int(ombfile["DAC Country Code"][i]))
                    if namibia == '275':  # Being used as a temporary stopgap
                        country = 'NA'
                    else:
                        country = '998'
                except ValueError:
                    country = '998'
        if country not in actives:
            actives.append(country)
            ids.append([country, i])
        else:
            for each in ids:
                if country == each[0]:
                    each.append(i)
    return ids


def activities_loop(allids):
    """
    Find which activities are the main ones to set as hierarchy 1.
    :param allids: All of the identifiers, without the award numbers.
    :return rows: The index numbers corresponding to the rows.
    """
    rows = list()
    actives = list()
    for i in range(0, len(allids)):
        if allids[i] not in actives:
            actives.append(allids[i])
            rows.append(i)
    return rows


def related_loop(theomb, theombawards, iden):
    """
    Find which activities are related to the main ones.
    :param theomb: The list of hierarchy 1 activities.
    :param theombawards: The list of every row with their own awards.
    :param iden: The current hierarchy 1 activity to relate to.
    :return related: The index numbers of the related activities.
    """
    related = list()
    relateds = list()
    for i in range(0, len(theomb)):
        if iden == theomb[i] and theombawards[i] not in relateds:
            relateds.append(theombawards[i])
            related.append(i)
    return related


def trans_loop(idawarded, award):
    """
    Finds all transactions sharing the current activity's award ID.
    :param idawarded: The list of activities with the proper award num.
    :param award: The current award number to look for.
    :return transactions: Index numbers of the related transactions.
    """
    transactions = list()
    for i in range(0, len(idawarded)):
        if award == idawarded[i]:
            transactions.append(i)
    return transactions


def cluster_loop(clustercodes, transelement):
    """
    Creates SubElements for each appropriate cluster code attached to a transaction.
    :param clustercodes: The list of cluster codes attached to the transaction.
    :param transelement: The SubElement to add on to.
    :return: N/A
    """
    # TODO: Insert the appropriate vocab uri.
    vocaburi = ""
    clusterlist = clustercodes.split(';')
    for code in clusterlist:
        SubElement(transelement, 'sector', code=code, vocabulary="10", vocabulary_h_uri=vocaburi)


def historical_loop(hist_dict, cleanou, histcode, histfile):
    """
    Finds all transactions sharing the current activity's clean OU.
    :param hist_dict: The dictionary of activities with the proper clean OU.
    :param cleanou: The current clean OU to look for.
    :param histcode: The region or country code to match with.
    :param histfile: The historical data file.
    :return transaction: Index numbers of the related transactions.
    """
    hists = []
    temphist = []

    try:
        for i in range(0, len(hist_dict[cleanou])):
            if histfile["DAC Regional Code"][hist_dict[cleanou][i]] == histcode or\
                    histfile["ISO Alpha Code"][hist_dict[cleanou][i]] == histcode:
                temphist.append(str(histfile["Award Transaction Type"][hist_dict[cleanou][i]]))
                try:
                    histamount = float(histfile["Award Transaction Value"][hist_dict[cleanou][i]])
                    histvalue = str('{0:.2f}'.format(histamount))
                except ValueError:
                    histvalue = '0.00'
                temphist.append(histvalue)

                try:
                    hist_date = str(int(histfile["Award Transaction Date"][hist_dict[cleanou][i]]))
                    formattedhistdate = hist_date[0:4] + '-' + hist_date[4:6] + '-' + hist_date[6:8]
                    temphist.append(formattedhistdate)
                except ValueError:
                    temphist.append('')

                # DAC Sectors
                try:
                    dachistcode = str(int(histfile["DAC Purpose Code"][hist_dict[cleanou][i]]))
                except ValueError:
                    dachistcode = '0'
                temphist.append(dachistcode)

                hists.append(temphist)
                temphist = []
    except KeyError:
        return []

    return hists


def dictfiles(locfile, docfile, histfile, resfile):
    """
    Create dictionaries for faster access of the mapping files.
    :param locfile: the location file to use for mapping.
    :param docfile: the document file to use for mapping.
    :param histfile: the historical data file to use for mapping.
    :param resfile: the results data file to use for mapping.
    :return loc_dict: the dict of locations
    :return doc_dict: the dict of documents
    :return hist_dict: the dict of historical data
    :return res_dict: the dict of the results and objectives data
    """
    loc_dict = {}
    doc_dict = {}
    hist_dict = {}
    res_dict = {}

    for i in range(0, len(locfile.index)):
        if locfile["clean_id"][i] in loc_dict:
            loc_dict[locfile["clean_id"][i]].append(i)
        else:
            loc_dict[locfile["clean_id"][i]] = [i]
    for j in range(0, len(docfile.index)):
        if docfile["clean_id"][j] in doc_dict:
            doc_dict[docfile["clean_id"][j]].append(j)
        else:
            doc_dict[docfile["clean_id"][j]] = [j]
    for h in range(0, len(histfile.index)):
        if histfile["Implementing Mechanism ID"][h] in hist_dict:
            hist_dict[histfile["Implementing Mechanism ID"][h]].append(h)
        else:
            hist_dict[histfile["Implementing Mechanism ID"][h]] = [h]
    for r in range(0, len(resfile.index)):
        if resfile["Clean ID"][r] in res_dict:
            res_dict[resfile["Clean ID"][r]].append(r)
        else:
            res_dict[resfile["Clean ID"][r]] = [r]
    print(loc_dict)
    print(doc_dict)
    print(hist_dict)
    print(res_dict)
    return loc_dict, doc_dict, hist_dict, res_dict


def location_loop(locfile, cleanid, locmap, loc_dict):
    """
    Create a list of sub-national locations for an activity.
    :param locfile: the location file to use for mapping.
    :param cleanid: the clean id of the award to map to.
    :param locmap: the ISO Alpha Code or DAC Regional Code of the award to map to.
    :param loc_dict: the dictionary holding the index to map to.
    :return locs: the list of locations to insert into the XML
    """
    locs = []
    temploc = []
    location_exact = "2"  # 1 is exact, 2 is approximate
    for i in range(0, len(loc_dict[cleanid])):
        if str(locfile["iso_alpha_code"][loc_dict[cleanid][i]]) == locmap \
                or str(locfile["dac_regional_code"][loc_dict[cleanid][i]]) == locmap:

            temploc.append(str(locfile["District"][loc_dict[cleanid][i]]))
            temploc.append(str(locfile["location_coordinates"][loc_dict[cleanid][i]]))
            try:
                temploc.append(str(int(locfile["location_reach"][loc_dict[cleanid][i]])))
            except ValueError:
                temploc.append('2')
            temploc.append(location_exact)
            temploc.append(str(int(locfile["location_type"][loc_dict[cleanid][i]])))
            locs.append(temploc)
            temploc = []
    return locs


def docs_loop(docfile, cleanid, doc_dict):
    """
    Create a list of documents for an activity.
    :param docfile: the documents file to use for mapping.
    :param cleanid: the clean id of the award to map to.
    :param doc_dict: the dictionary holding the index to map to.
    :return docs: the list of documents to insert into the XML
    """
    # Title, URL, Format, Category, Language
    docs = []
    tempdoc = []

    for i in range(0, len(doc_dict[cleanid])):
        tempdoc.append(str(docfile["Activity Title"][doc_dict[cleanid][i]]))
        tempdoc.append(str(docfile["file"][doc_dict[cleanid][i]]))
        tempdoc.append(str(docfile["doc_format"][doc_dict[cleanid][i]]))
        tempdoc.append(str(docfile["doc_category"][doc_dict[cleanid][i]]))
        tempdoc.append(str(docfile["Lang_code"][doc_dict[cleanid][i]]))

        try:
            doc_date = str(int(docfile["pubdate"][doc_dict[cleanid][i]]))
            formatteddocdate = doc_date[0:4] + '-' + doc_date[4:6] + '-' + doc_date[6:8]
            tempdoc.append(formatteddocdate)
        except ValueError:
            tempdoc.append('')

        docs.append(tempdoc)
        tempdoc = []

    return docs


def results_loop(resfile, cleanid, res_dict):
    """
    Create a list of documents for an activity.
    :param resfile: the results file to use for mapping.
    :param cleanid: the clean id of the award to map to.
    :param res_dict: the dictionary holding the index to map to.
    :return results: the list of results and objectives to insert into the XML
    """
    # Title, URL, Format, Category, Language
    resu = []
    tempres = []
    obje = []
    tempobj = []

    for i in range(0, len(res_dict[cleanid])):
        tempres.append(str(resfile["results"][res_dict[cleanid][i]]))
        tempres.append(str(resfile["results_title"][res_dict[cleanid][i]]))
        tempres.append(str(resfile["results_indicator"][res_dict[cleanid][i]]))
        tempobj.append(str(resfile["objectives"][res_dict[cleanid][i]]))

        resu.append(tempres)
        tempres = []
        obje.append(tempobj)
        tempobj = []
    return resu, obje


def lang_loop(transelement, langs, translations):
    """
    Create narratives for each language translation available.
    :param transelement: the main element to add the translation to.
    :param langs: the list of languages to translate into.
    :param translations: the list of translations.
    :return: N/A
    """
    it = 0
    for i in langs:
        if i is 'en':
            langnarrative = SubElement(transelement, 'narrative')
        else:
            langnarrative = SubElement(transelement, 'narrative', xml__lang=i)
        if translations[it] != 'nan':
            langnarrative.text = translations[it]
        else:
            langnarrative.text = 'There is no available description.'
        it += 1


def orgnumber(org):
    """
    Used to return the organization's registered IATI code.
    :param org: The organization to pull the code for
    :return: The organization's IATI code
    """
    if org == 'U.S. Government - U.S. Agency for International Development' \
            or org == 'U.S. Agency for International Development':
        orgcode = 'US-GOV-1'
    elif org == 'Department of Agriculture' or org == 'US Department of Agriculture' or org == 'Dept of Agriculture':
        orgcode = 'US-GOV-2'
    elif org == 'US Department of Treasury':
        orgcode = 'US-GOV-6'
    elif org == 'US Department of State' or org == 'State Department' or org == 'Dept of State':
        orgcode = 'US-GOV-11'
    elif org == 'Millennium Challenge Corporation' or org == 'MCC':
        orgcode = 'US-GOV-18'
    elif org == 'Exec Office of the President':
        orgcode = 'US-GOV-30'
    else:
        orgcode = ''
    return orgcode


# TODO: This currently has no use for us.
# TODO: We will need to evaluate what to do about percentages in the future.
def percentage_loop(reltrans, file):
    """
    Calcuate the percentage of the total money for an award a transaction consists of
    :param reltrans: The transaction to calculate the percentage for
    :param file: The omb file to pull data from
    :return alltotal: All of the money given towards an activity
    :return regiontotal: All of the money given towards a region/country
    :return sectortotal: All of the money given towards a specific sector
    """
    alltotal = 0
    regiontotal = list()
    sectortotal = list()
    countries = list()
    sectors = list()
    for i in reltrans:
        try:
            toadd = int(file["award_transaction_value"][i])
            alltotal += toadd
        except ValueError:
            toadd = 0

        try:
            country = str(int(file["dac_region_code"][i]))
        except ValueError:
            if str(file["iso_alpha_code"][i]) != 'nan':
                try:
                    country = str(file["iso_alpha_code"][i])
                except ValueError:
                    country = 'XXX'
            else:
                country = '998'
        if country not in countries:
            countries.append(country)
            regiontotal.append([country, toadd])
        else:
            for each in regiontotal:
                if country == each[0]:
                    each[1] += toadd

        try:
            sectorfind = str(int(file["us_sector_code"][i]))
            if sectorfind not in sectors:
                sectors.append(sectorfind)
                sectortotal.append([sectorfind, toadd])
            else:
                for each in sectortotal:
                    if sectorfind == each[0]:
                        each[1] += toadd
        except ValueError:
            pass
    return alltotal, regiontotal, sectortotal


def open_files():
    """
    Prompts the user for files to run the script on.
    :return omb: The omb file with the main data
    :return loc_file: The location file with the sub-national location mapping data
    :return doc_file: The documents file with the document mapping data
    :return hist_file: The historical data file with historical mapping data
    :return res_file: The results and objectives data file with mapping data
    """
    # Prompt user for filename
    # filetoopen = input("What is the name of the omb source file? ")
    filetoopen = 'FY18Q3/worldwide2.xlsx'
    print('Opening OMB file...')
    # Read the file
    try:
        ombf = pandas.read_excel(filetoopen, encoding='utf-8')
    except FileNotFoundError:
        sys.exit("OMB file does not exist.")
    # Output the number of rows
    print('Total rows: {0}'.format(len(ombf)))
    # See which headers are available
    print(list(ombf))

    # loctoopen = input("What is the name of the location mapping file? ")
    loctoopen = 'FY18Q3/Subnat mapping.xlsx'
    print('Opening location file...')
    # Read the file
    try:
        locs_file = pandas.read_excel(loctoopen, encoding='utf-8')
    except FileNotFoundError:
        sys.exit("Location file does not exist.")
    # Output the number of rows
    print('Total rows: {0}'.format(len(locs_file)))
    # See which headers are available
    print(list(locs_file))

    # doctoopen = input("What is the name of the document mapping file? ")
    doctoopen = 'FY18Q3/DEC mapping.xlsx'
    print('Opening document file...')
    # Read the file
    try:
        docs_file = pandas.read_excel(doctoopen, encoding='utf-8')
    except FileNotFoundError:
        sys.exit("Document file does not exist.")
    # Output the number of rows
    print('Total rows: {0}'.format(len(docs_file)))
    # See which headers are available
    print(list(docs_file))

    # histtoopen = input("What is the name of the historical data file? ")
    histtoopen = 'FY18Q3/historical_transactions.xlsx'
    print('Opening historical data file...')
    # Read the file
    try:
        hists_file = pandas.read_excel(histtoopen, encoding='utf-8')
    except FileNotFoundError:
        sys.exit("Historical file does not exist.")
    # Output the number of rows
    print('Total rows: {0}'.format(len(hists_file)))
    # See which headers are available
    print(list(hists_file))

    # restoopen = input("What is the name of the results data file? ")
    restoopen = 'FY18Q3/Obj Results mapping.xlsx'
    print('Opening results data file...')
    # Read the file
    try:
        resu_file = pandas.read_excel(restoopen, encoding='utf-8')
    except FileNotFoundError:
        sys.exit("Results file does not exist.")
    # Output the number of rows
    print('Total rows: {0}'.format(len(resu_file)))
    # See which headers are available
    print(list(resu_file))

    return ombf, locs_file, docs_file, hists_file, resu_file


curtime = time.time()
omb, loc_file, doc_file, hist_file, res_file = open_files()
opentime = time.time() - curtime
print('Converting format...')
now = datetime.datetime.utcnow().strftime('%Y-%m-%d')

# Variable creation
idlist, idawards, isolist = id_loop(omb)
locdict, docdict, histdict, resdict = dictfiles(loc_file, doc_file, hist_file, res_file)
# This will turn on full dataset dump into one XML. Untab all code after this.
# ombActs = activities_loop(idlist)

h1acts = activities_loop(idlist)

# This will turn on the splitting of the file via recipient if you uncomment this and tab everything after these.
ombgrouping = group_split(omb)
for ombActs in ombgrouping:

    filesleft = len(ombActs)

    ver = '2.03'
    fasite = 'https://explorer.usaid.gov/'

    activities = Element('iati-activities', version=ver,
                         generated_h_datetime=date, xmlns__usg=fasite)

    # Start creating the hierarchy 1 groupings
    c = 1
    # For loop for the amount of activities
    for act in ombActs:
        if c > 1:
            if act in h1acts:
                # Variables
                hier = '1'
                lastUpdate = date
                # TODO: Presumably, pull in data from alternate translation fields?
                langList = ['en']
                cur = 'USD'
                ident = idlist[act]
                repOrgRef = 'US-GOV-1'
                repOrgType = '10'  # Government
                repOrgText = 'U.S. Agency for International Development'
                partOrgText1 = str(omb["Appropriated Agency"][act])
                partOrgRef1 = orgnumber(partOrgText1)
                partOrgRole1 = '1'
                partOrgType1 = '10'  # Government
                partOrgRef2 = 'US-GOV-1'  # USAID
                partOrgRole2 = '2'
                partOrgText2 = 'U.S. Agency for International Development'
                partOrgType2 = '10'  # Government
                titleText = list()
                descText = list()

                # Create the activity group's name
                if str(omb["DAC Country Name"][act]) != 'nan':
                    try:
                        name = str(omb["DAC Country Name"][act])
                        if name == "CÃ´te dâ€™Ivoire":
                            name = "Côte d'Ivoire"
                        elif name == "Lao Peopleâ€™s Democratic Republic":
                            name = "Lao People's Democratic Republic"
                        title = 'US-' + name + '-' + partOrgText2
                    except ValueError:
                        name = str(omb["DAC Country Name"][act])
                        if name == "CÃ´te dâ€™Ivoire":
                            name = "Côte d'Ivoire"
                        elif name == "Lao Peopleâ€™s Democratic Republic":
                            name = "Lao People's Democratic Republic"
                        title = 'US-' + name + '-' + partOrgText2
                else:
                    title = 'US-Worldwide-' + partOrgText2
                titleText.append(title)
                titleText.append('')
                descText.append(str(omb["Implementing Mechanism Purpose Statement"][act]))
                descText.append('')

                relatedList = related_loop(idlist, idawards, ident)

                for rel in relatedList:
                    relType = '2'
                    relRef = idawards[rel]

                # Begin creating the activities
                for relact in relatedList:
                    clean_id = str(omb["Clean ID"][relact])
                    clean_ou = str(omb["Clean OU Name"][relact])
                    award_id = str(omb["Implementing Mechanism ID"][relact])
                    hier = '1'
                    lastUpdate = date
                    langList = ['en']
                    cur = 'USD'
                    countryinit = isolist[relact]
                    identity = idawards[relact]
                    repOrgRef = 'US-GOV-1'
                    repOrgType = '10'  # Government
                    repOrgText = 'U.S. Agency for International Development'
                    # These will need to be looped through and placed into narratives.
                    titleText = list()
                    descText = list()
                    titleText.append(str(omb["Implementing Mechanism Title"][relact]))
                    titleText.append('')
                    descText.append(str(omb["Implementing Mechanism Purpose Statement"][relact]))
                    descText.append('')
                    partOrgText = str(omb["Appropriated Agency"][relact])
                    partOrgRef = orgnumber(partOrgText)
                    partOrgRole = '1'
                    partOrgRef2 = 'US-GOV-1'
                    partOrgRole2 = '2'
                    partOrgText2 = 'U.S. Agency for International Development'
                    partOrgType2 = '10'
                    partOrgRole3 = '3'
                    partOrgText3 = 'U.S. Agency for International Development'
                    # These may change, depending on input.
                    # TODO: If more organizations become used, they will need impl.
                    partOrgRef3 = orgnumber(partOrgText3)
                    partOrgType3 = '10'
                    partOrgRole4 = '4'
                    partOrgText4 = str(omb["Implementing Agent"][relact])
                    if str(omb["IATI Organization ID"][relact]) == 'nan':
                        partOrgRef4 = orgnumber(partOrgText4)
                    else:
                        partOrgRef4 = str(omb["IATI Organization ID"][relact])

                    try:
                        partOrgType4 = str(int(omb["Implementing Agent Type"][relact]))
                    except ValueError:
                        partOrgType4 = ''
                    try:
                        activityStatusCode = str(int(omb["Reporting Status"][relact]))
                    except ValueError:
                        activityStatusCode = '1'

                    # All dates are always "actual". There are no "planned" dates.
                    try:
                        isodatetime = str(int(omb["Start Date"][relact]))
                        isodatetimeformatstart = isodatetime[0:4] + '-' +\
                            isodatetime[4:6] + '-' + isodatetime[6:8]
                        activityStartDateText = ''
                    except ValueError:
                        isodatetimeformatstart = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
                        activityStartDateText = str(omb["start_date_narr"][relact])
                    try:
                        isodatetime = str(int(omb["End Date"][relact]))
                        isodatetimeformatend = isodatetime[0:4] + '-' + isodatetime[4:6] +\
                            '-' + isodatetime[6:8]
                        activityEndDateText = ''
                    except ValueError:
                        isodatetimeformatend = datetime.datetime.utcnow().strftime('%Y') + '-10-01'
                        activityEndDateText = str(omb["end_date_narr"][relact])
                    activityDateTypePlanStart = '1'
                    activityDateTypeStart = '2'
                    activityDateTypePlanEnd = '3'
                    activityDateTypeEnd = '4'
                    activityScopeCode = str(int(omb["Activity Scope"][relact]))
                    try:
                        signdate = \
                            str(int(omb["Implementing Mechanism Signing Date"][relact]))
                        signdateformat = signdate[0:4] + '-' + signdate[4:6] + '-' +\
                            signdate[6:8]
                    except ValueError:
                        signdateformat = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'

                    # Put together the first part of the activity element tree
                    activity = SubElement(activities, 'iati-activity', hierarchy=hier,
                                          last_h_updated_h_datetime=lastUpdate,
                                          xml__lang=langList[0], default_h_currency=cur)
                    identifier = SubElement(activity, 'iati-identifier')
                    repId = repOrgRef + '-' + award_id
                    identifier.text = repId
                    reporting_org = SubElement(activity, 'reporting-org', ref=repOrgRef,
                                               type=repOrgType)
                    narrative = SubElement(reporting_org, 'narrative')
                    narrative.text = repOrgText
                    title = SubElement(activity, 'title')
                    lang_loop(title, langList, titleText)
                    description = SubElement(activity, 'description')
                    lang_loop(description, langList, descText)

                    # Populate the results and objectives
                    resList = []
                    objList = []
                    if clean_id != 'nan':
                        if clean_id in resdict:
                            resList, objList = results_loop(res_file, clean_id, resdict)

                        for obj in objList:
                            if obj[0] != 'nan' and obj[0] != '':
                                transObjective = SubElement(activity, 'description', type='2')
                                narrative = SubElement(transObjective, 'narrative')
                                narrative.text = obj[0]

                    participating_org1 = SubElement(activity, 'participating-org',
                                                    ref=partOrgRef, role=partOrgRole, type=partOrgType1)
                    narrative = SubElement(participating_org1, 'narrative')
                    narrative.text = partOrgText1
                    participating_org2 = SubElement(activity, 'participating-org',
                                                    ref=partOrgRef2, role=partOrgRole2, type=partOrgType2)
                    narrative = SubElement(participating_org2, 'narrative')
                    narrative.text = partOrgText2
                    participating_org3 = SubElement(activity, 'participating-org',
                                                    ref=partOrgRef3, role=partOrgRole3, type=partOrgType3)
                    narrative = SubElement(participating_org3, 'narrative')
                    narrative.text = partOrgText3
                    if partOrgRef4 != '':
                        if partOrgType4 != 'nan':
                            participating_org4 = SubElement(activity, 'participating-org',
                                                            ref=partOrgRef4, role=partOrgRole4, type=partOrgType4)
                        else:
                            participating_org4 = SubElement(activity, 'participating-org',
                                                            ref=partOrgRef4, role=partOrgRole4)
                    else:
                        if partOrgType4 != 'nan':
                            participating_org4 = SubElement(activity, 'participating-org',
                                                            role=partOrgRole4, type=partOrgType4)
                        else:
                            participating_org4 = SubElement(activity, 'participating-org',
                                                            role=partOrgRole4)

                    narrative = SubElement(participating_org4, 'narrative')
                    if partOrgText4 != 'nan':
                        narrative.text = partOrgText4
                    else:
                        narrative.text = '--'

                    activity_status = SubElement(activity, 'activity-status',
                                                 code=activityStatusCode)

                    # All dates are always "actual". There are no "planned" dates.
                    # activity_planstart = SubElement(activity, 'activity-date',
                    #                            iso_h_date=isodatetimeformatstart,
                    #                            type=activityDateTypePlanStart)
                    activity_planstartdate = SubElement(activity, 'activity-date',
                                                        iso_h_date=isodatetimeformatstart,
                                                        type=activityDateTypePlanStart)

                    if isodatetimeformatstart <= now:
                        activity_startdate = SubElement(activity, 'activity-date',
                                                        iso_h_date=isodatetimeformatstart, type=activityDateTypeStart)
                    if activityStartDateText:
                        narrative = SubElement(activity_planstartdate, 'narrative')
                        narrative.text = activityStartDateText

                    # All dates are always "actual". There are no "planned" dates.
                    # activity_planend = SubElement(activity, 'activity-date',
                    #                               iso_h_date=isodatetimeformatend,
                    #                               type=activityDateTypePlanEnd)
                    activity_planenddate = SubElement(activity, 'activity-date',
                                                      iso_h_date=isodatetimeformatend, type=activityDateTypePlanEnd)
                    if isodatetimeformatend <= now:
                        activity_enddate = SubElement(activity, 'activity-date',
                                                      iso_h_date=isodatetimeformatend, type=activityDateTypeEnd)
                    if activityEndDateText:
                        narrative = SubElement(activity_planenddate, 'narrative')
                        narrative.text = activityEndDateText

                    # Contact information block variables
                    contactType = '1'
                    organisationText = 'U.S. Agency for International Development'
                    personNameText = str(omb["USAID contact name"][relact])
                    telephoneText = str(omb["USAID contact telephone"][relact])
                    emailText = str(omb["USAID contact email"][relact])
                    websiteText = str(omb["Activity Website"][relact])
                    mailingText = str(omb["USAID contact address"][relact])

                    # Contact information block
                    contact_info = SubElement(activity, 'contact-info', type=contactType)
                    organisation = SubElement(contact_info, 'organisation')
                    narrative = SubElement(organisation, 'narrative')
                    narrative.text = organisationText
                    person_name = SubElement(contact_info, 'person-name')
                    narrative = SubElement(person_name, 'narrative')
                    if personNameText != 'nan':
                        narrative.text = personNameText
                    telephone = SubElement(contact_info, 'telephone')
                    if telephoneText != 'nan':
                        telephone.text = telephoneText
                    email = SubElement(contact_info, 'email')
                    if emailText != 'nan':
                        email.text = emailText
                    website = SubElement(contact_info, 'website')
                    if websiteText != 'nan':
                        website.text = websiteText
                    mailing_address = SubElement(contact_info, 'mailing-address')
                    narrative = SubElement(mailing_address, 'narrative')
                    if mailingText != 'nan':
                        narrative.text = mailingText

                    activity_scope = SubElement(activity, 'activity-scope',
                                                code=activityScopeCode)

                    # Variables
                    recipientCountryPercentage = '100'  # This will never not be 100%

                    # Pre-transaction information block
                    if str(omb["ISO Alpha Code"][relact]) == 'nan':
                        recipientRegionCode = countryinit
                        recipient_region = SubElement(activity, 'recipient-region',
                                                      percentage=recipientCountryPercentage,
                                                      code=recipientRegionCode)
                    else:
                        recipientCountryCode = countryinit
                        recipient_country = SubElement(activity, 'recipient-country',
                                                       percentage=recipientCountryPercentage,
                                                       code=recipientCountryCode)

                    # Populate the subnational locations
                    if clean_id != 'nan':
                        if clean_id in locdict:
                            locsList = location_loop(loc_file, clean_id, countryinit, locdict)
                        else:
                            locsList = []
                        gis = "http://www.opengis.net/def/crs/EPSG/0/4326"
                        for loc in locsList:
                            # FIX: url and format get flipped somehow?
                            location = SubElement(activity, 'location')
                            reach = SubElement(location, 'location-reach', code=loc[2])
                            name = SubElement(location, 'name')
                            narrative = SubElement(name, 'narrative')
                            narrative.text = loc[0]
                            point = SubElement(location, 'point', srsName=gis)
                            pos = SubElement(point, 'pos')
                            pos.text = loc[1]
                            exactness = SubElement(location, 'exactness', code=loc[3])
                            locationclass = SubElement(location, 'location-class', code=loc[4])

                    try:
                        collabCode = str(int(omb["Collaboration Type Code"][relact]))
                    except ValueError:
                        if str(omb["Collaboration Type"][relact]) == 'Bilateral':
                            collabCode = '1'
                        else:
                            collabCode = '2'
                    # collabCode = str(int(omb["Collaboration Type Code"][relact]))
                    try:
                        flowType = str(int(omb["Flow Type"][relact]))
                    except ValueError:
                        flowType = '0'
                    try:
                        financeType = str(int(omb["Finance Type"][relact]))
                    except ValueError:
                        financeType = '0'
                    try:
                        aidType = str(omb["Aid Type Code"][relact])
                    except ValueError:
                        aidType = '0'
                    try:
                        tiedCode = str(int(omb["Tying Status of Award"][relact]))
                    except ValueError:
                        tiedCode = '0'
                    # This is for error checking
                    periodStartDate = ''
                    try:
                        if str(int(omb["Start Date"][relact])) != 'nan':
                            periodStart = \
                                str(int(omb["Start Date"][relact]))
                            periodStartDate = periodStart[0:4] + '-' + periodStart[4:6] + '-' + periodStart[6:8]
                    except ValueError:
                        periodStartDate = str(int(omb["Beginning Fiscal Funding Year"][relact])) + '-10-01'
                    periodEndDate = ''
                    try:
                        if str(int(omb["End Date"][relact])) != 'nan':
                            periodEnd = \
                                str(int(omb["End Date"][relact]))
                            periodEndDate = periodEnd[0:4] + '-' + periodEnd[4:6] + '-' + periodEnd[6:8]
                    except ValueError:
                        try:
                            if str(int(omb["Ending Fiscal Funding Year"][relact])) != 'nan':
                                periodEndDate = str(int(omb["Ending Fiscal Funding Year"][relact])) + '-09-30'
                        except ValueError:
                            periodEndDate = ''
                    budgetValueDate = periodStartDate
                    try:
                        totalallocationsfloat = float(omb["Total allocations"][relact])
                        budgetAmount = str('{0:.2f}'.format(omb["Total allocations"][relact]))
                        if budgetAmount == 'nan':
                            budgetAmount = '0.00'
                    except ValueError:
                        budgetAmount = '0.00'
                    budgetStatus = "1"

                    # Create the pre-transaction types element tree
                    collaboration_type = SubElement(activity, 'collaboration-type',
                                                    code=collabCode)
                    if flowType != '0':
                        default_flow_type = SubElement(activity, 'default-flow-type',
                                                       code=flowType)
                    if financeType != '0':
                        default_finance_type = SubElement(activity, 'default-finance-type',
                                                          code=financeType)
                    if aidType != '0':
                        default_aid_type = SubElement(activity, 'default-aid-type',
                                                      code=aidType)
                    if tiedCode != '0':
                        default_tied_status = SubElement(activity, 'default-tied-status',
                                                         code=tiedCode)

                    # Budget block
                    if budgetAmount != '0.00':
                        budget = SubElement(activity, 'budget', status=budgetStatus)
                        SubElement(budget, 'period-start', iso_h_date=periodStartDate)
                        SubElement(budget, 'period-end', iso_h_date=periodEndDate)
                        budgetValue = SubElement(budget, 'value', currency=cur,
                                                 value_h_date=budgetValueDate)
                        budgetValue.text = budgetAmount

                    # Create the list of transactions for a specific activity
                    # histList = trans_loop(histdict, idawards[relact])
                    histList = historical_loop(histdict, award_id, countryinit, hist_file)
                    transList = trans_loop(idawards, idawards[relact])

                    # These two make it easier to determine
                    # if a 0 has been put in for com/dis
                    # If changed to True, then a 0 transaction will not be put into XML
                    comMarker = True
                    disMarker = True

                    # Loop through the historical transactions
                    for trans in histList:

                        transaction = ''
                        valueAmount = trans[1]
                        transType = trans[0]
                        if transType == "Commitment" or transType == "Obligation":
                            transaction_code = '2'
                        elif transType == "Disbursement":
                            transaction_code = '3'
                        else:
                            transaction_code = '0'
                        value_datetime = trans[2]
                        # Make sure there is exactly 1 transaction value of 0.00 for Com if needed
                        if transaction_code == '2':
                            if (comMarker is True and valueAmount != '0.00')\
                                    or (comMarker is False and valueAmount == '0.00'):
                                # Set the elements
                                transaction = SubElement(activity, 'transaction')
                                transaction_type = SubElement(transaction, 'transaction-type',
                                                              code=transaction_code)
                                transaction_date = SubElement(transaction, 'transaction-date',
                                                              iso_h_date=value_datetime)
                                value = SubElement(transaction, 'value',
                                                   value_h_date=value_datetime)
                                value.text = valueAmount

                            if str(int(trans[3])) != '0':
                                sector = SubElement(transaction, 'sector',
                                                    code=str(int(trans[3])), vocabulary='1')
                            # Cluster Codes
                            # TODO: adjust the cluster code column name
                            try:
                                clusters = ""
                                # clusters = str(omb["cluster_codes"][trans])
                            except ValueError:
                                clusters = ""

                            if clusters != "" and clusters != "nan":
                                cluster_loop(clusters, transaction)

                        # Make sure there is exactly 1 transaction value of 0.00 for Disb if needed
                        if transaction_code == '3':
                            if (disMarker is True and valueAmount != 0) or (
                                    disMarker is False and valueAmount == 0):
                                # Set the elements
                                transaction = SubElement(activity, 'transaction')
                                transaction_type = SubElement(transaction, 'transaction-type',
                                                              code=transaction_code)
                                transaction_date = SubElement(transaction, 'transaction-date',
                                                              iso_h_date=value_datetime)
                                value = SubElement(transaction, 'value',
                                                   value_h_date=value_datetime)
                                value.text = valueAmount

                            # DAC Sectors
                            if str(int(trans[3])) != '0':
                                sector = SubElement(transaction, 'sector',
                                                    code=str(int(trans[3])), vocabulary='1')
                            # Cluster Codes
                            # TODO: adjust the cluster code column name
                            try:
                                clusters = ""
                                # clusters = str(omb["cluster_codes"][trans])
                            except ValueError:
                                clusters = ""

                            if clusters != "" and clusters != "nan":
                                cluster_loop(clusters, transaction)

                    # Loop through the transactions related to the activity
                    for trans in transList:
                        # Variables that depend on entries
                        # If the disbursement has a value, set value to disbursement.
                        try:
                            transAmount = float(omb["Award Transaction Value"][trans])
                            valueAmount = str('{0:.2f}'.format(transAmount))
                        except ValueError:
                            valueAmount = '0.00'
                        transDescList = list()
                        transDescList.append(str(omb["Award Transaction - Description"][trans]))
                        transDescList.append('')
                        transType = str(omb["Award Transaction Type"][trans])
                        if transType == "Commitment" or transType == "Obligation":
                            transaction_code = '2'
                        elif transType == "Disbursement":
                            transaction_code = '3'
                        else:
                            transaction_code = '0'
                        try:
                            valuedate = str(int(omb["Award Transaction Date"][trans]))
                            value_datetime = valuedate[0:4] + '-' + valuedate[4:6] + \
                                '-' + valuedate[6:8]
                        except ValueError:
                            value_datetime = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
                        regAccCode = str(int(omb["Treasury Regular Account Code"][trans]))
                        mainAccCode = str(int(omb["Treasury Main Account Code"][trans]))
                        mainText = str(omb["Treasury Main Account Title"][trans])
                        try:
                            fundingYearBegin = \
                                str(int(omb["Beginning Fiscal Funding Year"][trans]))
                        except ValueError:
                            fundingYearBegin = str(int(datetime.datetime.utcnow().strftime('%Y'))-1)
                        try:
                            fundingYearEnd = \
                                str(int(omb["Ending Fiscal Funding Year"][trans]))
                        except ValueError:
                            fundingYearEnd = str(datetime.datetime.utcnow().strftime('%Y'))

                        transaction = ''
                        # Make sure there is exactly 1 transaction value of 0.00 for Com if needed
                        if transaction_code == '2':
                            if (comMarker is True and valueAmount != '0.00') or\
                                    (comMarker is False and valueAmount == '0.00'):
                                # Set the elements
                                try:
                                    if str(int(omb["Humanitarian Tag"][trans])) == '1':
                                        transaction = SubElement(activity, 'transaction', humanitarian='1')
                                    else:
                                        transaction = SubElement(activity, 'transaction')
                                except ValueError:
                                    transaction = SubElement(activity, 'transaction')
                                transaction_type = SubElement(transaction, 'transaction-type',
                                                              code=transaction_code)
                                transaction_date = SubElement(transaction, 'transaction-date',
                                                              iso_h_date=value_datetime)
                                value = SubElement(transaction, 'value',
                                                   value_h_date=value_datetime)
                                value.text = valueAmount
                                transDescription = SubElement(transaction, 'description')
                                lang_loop(transDescription, langList, transDescList)

                                try:
                                    disbChan = str(int(omb["Disbursement Channel"][trans]))
                                except ValueError:
                                    disbChan = '0'

                                # Sectors
                                try:
                                    dacCode = str(int(omb["DAC Purpose Code"][trans]))
                                except ValueError:
                                    dacCode = '0'
                                try:
                                    sectorCode = str(int(omb["U.S. Government Sector Code"][trans]))
                                except ValueError:
                                    sectorCode = '0'
                                dacVocab = '1'
                                dacText = str(omb["DAC Purpose Name"][trans])
                                sectorVocab = '99'
                                sectorText = str(omb["U.S. Government Sector Name"][trans])

                                # Create the element tree
                                disburseChannel = SubElement(transaction, 'disbursement-channel', code=disbChan)
                                if dacCode != '0':
                                    sector = SubElement(transaction, 'sector',
                                                        code=dacCode, vocabulary=dacVocab)
                                    narrative = SubElement(sector, 'narrative')
                                    narrative.text = dacText
                                sector = SubElement(transaction, 'sector',
                                                    code=sectorCode, vocabulary=sectorVocab)
                                narrative = SubElement(sector, 'narrative')
                                narrative.text = sectorText

                                try:
                                    if str(int(omb["Humanitarian Tag"][trans])) == '1':
                                        cluster = str(int(omb["Cluster ID"][trans]))
                                        if cluster != 'nan':
                                            clusterSector = SubElement(transaction, 'sector',
                                                                       code=cluster, vocabulary='10')
                                except ValueError:
                                    cluster = '0'

                                treasury_account = \
                                    SubElement(transaction, 'usg__treasury-account')
                                regular_account = SubElement(treasury_account,
                                                             'usg__regular-account',
                                                             code=regAccCode)
                                main_account = SubElement(treasury_account, 'usg__main-account',
                                                          code=mainAccCode)
                                main_account.text = mainText
                                fiscal_funding_year = SubElement(treasury_account,
                                                                 'usg__fiscal-funding-year',
                                                                 begin=fundingYearBegin,
                                                                 end=fundingYearEnd)
                            comMarker = True

                        # Make sure there is exactly 1 transaction value of 0.00 for Disb if needed
                        if transaction_code == '3':
                            if (disMarker is True and valueAmount != 0) or (disMarker is False and valueAmount == 0):
                                # Set the elements
                                try:
                                    if str(int(omb["Humanitarian Tag"][trans])) == '1':
                                        transaction = SubElement(activity, 'transaction', humanitarian='1')
                                    else:
                                        transaction = SubElement(activity, 'transaction')
                                except ValueError:
                                    transaction = SubElement(activity, 'transaction')
                                transaction_type = SubElement(transaction, 'transaction-type',
                                                              code=transaction_code)
                                transaction_date = SubElement(transaction, 'transaction-date',
                                                              iso_h_date=value_datetime)
                                value = SubElement(transaction, 'value',
                                                   value_h_date=value_datetime)
                                value.text = valueAmount
                                transDescription = SubElement(transaction, 'description')
                                lang_loop(transDescription, langList, transDescList)

                                # TODO: Adjust objective description so that only one shows up in an activity
                                # if actObj != 'nan' and actObj != '':
                                #    transObjective = SubElement(transaction, 'description', type='2')
                                #    narrative = SubElement(transObjective, 'narrative')
                                #    narrative.text = str(omb["Activity Objective"][trans])

                                try:
                                    disbChan = str(int(omb["Disbursement Channel"][trans]))
                                except ValueError:
                                    disbChan = '0'

                                # Sectors
                                try:
                                    dacCode = str(int(omb["DAC Purpose Code"][trans]))
                                except ValueError:
                                    dacCode = '0'
                                try:
                                    sectorCode = str(int(omb["U.S. Government Sector Code"][trans]))
                                except ValueError:
                                    sectorCode = '0'
                                dacVocab = '1'
                                dacText = str(omb["DAC Purpose Name"][trans])
                                sectorVocab = '99'
                                sectorText = str(omb["U.S. Government Sector Name"][trans])

                                try:
                                    if str(int(omb["Humanitarian Tag"][trans])) == '1':
                                        cluster = str(int(omb["Cluster ID"][trans]))
                                        if cluster != 'nan':
                                            clusterSector = SubElement(transaction, 'sector',
                                                                       code=cluster, vocabulary='10')
                                except ValueError:
                                    cluster = '0'

                                # Create the element tree
                                disburseChannel = SubElement(transaction, 'disbursement-channel', code=disbChan)
                                if dacCode != '0':
                                    sector = SubElement(transaction, 'sector',
                                                        code=dacCode, vocabulary=dacVocab)
                                    narrative = SubElement(sector, 'narrative')
                                    narrative.text = dacText
                                sector = SubElement(transaction, 'sector',
                                                    code=sectorCode, vocabulary=sectorVocab)
                                narrative = SubElement(sector, 'narrative')
                                narrative.text = sectorText

                                treasury_account = \
                                    SubElement(transaction, 'usg__treasury-account')
                                regular_account = SubElement(treasury_account,
                                                             'usg__regular-account',
                                                             code=regAccCode)
                                main_account = SubElement(treasury_account, 'usg__main-account',
                                                          code=mainAccCode)
                                main_account.text = mainText
                                fiscal_funding_year = SubElement(treasury_account,
                                                                 'usg__fiscal-funding-year',
                                                                 begin=fundingYearBegin,
                                                                 end=fundingYearEnd)
                            disMarker = True

                    # Populate the document links
                    if clean_id != 'nan':
                        if clean_id in docdict:
                            docsList = docs_loop(doc_file, clean_id, docdict)
                        else:
                            docsList = []
                        for doc in docsList:
                            # FIX: url and format get flipped somehow?
                            document = SubElement(activity, 'document-link', format=doc[2], url=doc[1])
                            title = SubElement(document, 'title')
                            narrative = SubElement(title, 'narrative')
                            narrative.text = doc[0]
                            category = SubElement(document, 'category', code=doc[3])
                            lang = SubElement(document, 'language', code=doc[4])
                            if doc[5] != '':
                                docdate = SubElement(document, 'document-date', iso_h_date=doc[5])
                    # conditionsDocument = str(omb["Conditions Document Link"][relact])
                    # if conditionsDocument != 'nan':
                    #     if conditionsDocument == "https://www.usaid.gov/sites/default/files/documents/1868/302.pdf":
                    #         conditionsDocumentTitle = "ADS Chapter 302 USAID Direct Contracting"
                    #     elif conditionsDocument == "https://www.usaid.gov/sites/default/files/documents/1868/303.pdf":
                    #         conditionsDocumentTitle = \
                    #             "ADS Chapter 303 Grants and Cooperative Agreements to Non-Governmental Organizations"
                    #     conditionsAttached = "1"
                    #     document = SubElement(activity, 'document-link', format="application/pdf",
                    #                           url=conditionsDocument)
                    #     title = SubElement(document, 'title')
                    #     narrative = SubElement(title, 'narrative')
                    #     narrative.text = conditionsDocumentTitle
                    #     category = SubElement(document, 'category', code="A04")
                    #     lang = SubElement(document, 'language', code="en")
                    # else:
                    conditionsAttached = "0"

                    # TODO: Insert code for Contract Links here
                    # document-link code=A11
                    # Concatenate "https://www.usaspending.gov/Pages/AdvancedSearch.aspx?k=" + field
                    # if str(omb["stripped award"][relact]) != nan:
                    #  contractlink = (link) + str(omb["stripped award"][relact])
                    #  contract = subelement(activity, 'document-link', format=html, url=contractlink)
                    #  subelement(contract, 'category', code="A11")
                    #  subelement(contract, 'language', code="en")

                    # This assumes that there will never be any conditions.
                    # This is currently the case, however, this may eventually change.
                    conditions = SubElement(activity, 'conditions', attached=conditionsAttached)
                    SubElement(activity, 'usg__mechanism-signing-date',
                               iso_h_date=signdateformat)

                    for res in resList:
                        if res[1] != 'nan':
                            resulting = SubElement(activity, 'result', type='9')
                            resulttitle = SubElement(resulting, 'title')
                            narrative = SubElement(resulttitle, 'narrative')
                            narrative.text = res[1]
                            if res[0] != 'nan':
                                resultdescription = SubElement(resulting, 'description')
                                narrative = SubElement(resultdescription, 'narrative')
                                narrative.text = res[0]
                            if res[2] != 'nan':
                                resultindicator = SubElement(resulting, 'indicator', measure='5')
                                indicatortitle = SubElement(resultindicator, 'title')
                                narrative = SubElement(indicatortitle, 'narrative')
                                narrative.text = res[2]

                    # Extra fields requested by State
                    try:
                        duns = str(int(omb["Implementing Agent's DUNS Number"][relact]))
                    except ValueError:
                        duns = 'nan'
                    tec = str('{0:.2f}'.format(float(omb["TEC"][relact])))
                    stateloc = str(omb["State Location"][relact])
                    if stateloc == "CÃ´te d'Ivoire":
                        stateloc = "Côte d'Ivoire"
                    elif stateloc == "Lao Peopleâ€™s Democratic Republic":
                        stateloc = "Lao People's Democratic Republic"

                    if duns != 'nan':
                        dunselement = SubElement(activity, 'usg__duns-number')
                        narrative = SubElement(dunselement, 'narrative')
                        narrative.text = duns
                    if tec != 'nan':
                        tecelement = SubElement(activity, 'usg__tec1')
                        narrative = SubElement(tecelement, 'narrative')
                        narrative.text = tec
                    if stateloc != 'nan':
                        stateelement = SubElement(activity, 'usg__state-location')
                        narrative = SubElement(stateelement, 'narrative')
                        narrative.text = stateloc

        c += 1

    # End of run processing and time keeping stats.
    converttime = time.time() - curtime
    print('Writing file...')

    # This is to write to a singular file.
    # output_file = open('iati-activities-full.xml', 'w', encoding='utf-8')

    if not os.path.exists('export/' + time.strftime("%m-%d-%Y") + '/'):
        os.makedirs('export/' + time.strftime("%m-%d-%Y") + '/')
    # This line is for country names
    # TODO: int(ombActs[1]) <-> int(act)
    output_file = open('export/' + time.strftime("%m-%d-%Y")+'/iati-activities-Worldwide 2.xml',
                       'w', encoding='utf-8')
    # output_file = open('export/' + time.strftime("%m-%d-%Y") + '/iati-activities-' +
    #                   str(omb["Country File Name"][int(ombActs[1])])+'.xml', 'w', encoding='utf-8')
    # This line is for country codes
    # output_file = open('export/' + time.strftime("%m-%d-%Y") +
    # '/iati-activities-'+ombActs[0]+'.xml', 'w', encoding='utf-8')
    output_file.write(prettify(activities).replace("__", ":").replace("_h_", "-"))
    output_file.close()
    finaltime = time.time() - curtime
    print('Opening Time: ' + str(opentime))
    print('Convert Time: ' + str(converttime - opentime))
    print('Write Time: ' + str(finaltime - converttime))
    print('Run time: ' + str(finaltime))
    print('Average time per main activity: ' +
          str((converttime - opentime)/len(ombActs)))
    # print('Files left: ' + str(len(ombActs)))
print('Zipping...')
shutil.make_archive('export/zip/export-'+time.strftime("%m-%d-%Y"), 'zip', 'export/' + time.strftime("%m-%d-%Y") + '/')
print('Complete!')
