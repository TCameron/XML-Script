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
__date__ = "05-24-2016"
__version__ = "0.21"
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
    # USAID is identified as US-GOV-1 within PWYF.
    code = '1'
    for i in range(0, len(ombfile.index)):
        try:
            country = str(int(ombfile["dac_regional_code"][i]))
        except ValueError:
            if str(ombfile["iso_alpha_code"][i]) != 'nan':
                try:
                    country = str(ombfile["iso_alpha_code"][i])
                except ValueError:
                    country = '998'
            else:
                # This is due to Namibia's code being "NA", but numpy counts "NA" as not applicable, so it gets skipped
                try:
                    namibia = str(int(ombfile["dac_country_code"][i]))
                    if namibia == '275':  # Being used as a temporary stopgap
                        country = 'NA'
                    else:
                        country = '998'
                except ValueError:
                    country = '998'
        try:
            categorytype = str(int(omb["us_sector_code"][i]))[0:2]
        except ValueError:
            categorytype = '90'
        awardid = str(omb["implementing_mechanism_id"][i])
        entry = 'US-GOV' + '-' + code + '-' + country + '-' + categorytype
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
            country = str(int(ombfile["dac_regional_code"][i]))
        except ValueError:
            print(str(ombfile["iso_alpha_code"][i]))
            if str(ombfile["iso_alpha_code"][i]) != 'nan':
                try:
                    country = str(ombfile["iso_alpha_code"][i])
                except ValueError:
                    country = '998'
            else:
                try:
                    namibia = str(int(ombfile["dac_country_code"][i]))
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


def location_loop(ombfile, award):
    """
    Create a list of sub-national locations for an activity.
    :param ombfile: the ombfile to pull data from.
    :param award: the award to find the list of locations for.
    :return locs: the list of locations to insert into the XML
    """
    locs = []
    temploc = []
    location_reach = "2"  # 1 is Activity, 2 is Intended Beneficiaries
    location_exact = "2"  # 1 is exact, 2 is approximate
    location_class = "1"  # 1 is Admin Region, 2 is Pop Place, 3 is Structure, 4 is Other
    i = 1
    while i:
        try:
            if str(ombfile["loc_name"+str(i)][award]) != 'nan':
                temploc.append(str(ombfile["loc_name"+str(i)][award]))
                temploc.append(str(ombfile["loc_coor"+str(i)][award]))
                temploc.append(location_reach)
                temploc.append(location_exact)
                temploc.append(location_class)
                locs.append(temploc)
                temploc = []
                i += 1
            else:
                return locs
        except:
            return locs
    return locs


def docs_loop(ombfile, award):
    """
    Create a list of documents for an activity.
    :param ombfile: the ombfile to pull data from.
    :param award: the award to find the list of documents for.
    :return docs: the list of documents to insert into the XML
    """
    # Title, URL, Format, Category, Language
    docs = []
    tempdoc = []
    i = 1
    j = 1
    while i:
        try:
            if str(ombfile["eval_title" + str(i)][award]) != 'nan':
                tempdoc.append(str(ombfile["eval_title" + str(i)][award]))
                tempdoc.append(str(ombfile["eval_link" + str(i)][award]))
                tempdoc.append(str(ombfile["eval_file_format"][award]))
                tempdoc.append(str(ombfile["eval_category"][award]))
                tempdoc.append(str(ombfile["eval_lang" + str(i)][award]))
                docs.append(tempdoc)
                tempdoc = []
                i += 1
            else:
                i = False
        except:
            i = False
    while j:
        try:
            if str(ombfile["imp_title" + str(j)][award]) != 'nan':
                tempdoc.append(str(ombfile["imp_title" + str(j)][award]))
                tempdoc.append(str(ombfile["imp_link" + str(j)][award]))
                tempdoc.append(str(ombfile["imp_file_format"][award]))
                tempdoc.append(str(ombfile["imp_category"][award]))
                tempdoc.append(str(ombfile["imp_lang" + str(j)][award]))
                docs.append(tempdoc)
                tempdoc = []
                j += 1
            else:
                j = False
        except:
            j = False
    return docs


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
            narrative = SubElement(transelement, 'narrative')
        else:
            narrative = SubElement(transelement, 'narrative', xml__lang=i)
        if translations[it] != 'nan':
            narrative.text = translations[it]
        else:
            narrative.text = 'There is no available description.'
        it += 1


def catdesc(cat):
    """
    Used to create the description for a hierarchy 1 grouping.
    :param cat: The category of the group
    :return: The description for the group, as provided by State.
    """
    if cat == '10':
        catdescription = 'To help nations effectively establish the conditions and capacity for achieving peace, ' \
                         'security, and stability; and for responding effectively against arising threats to national' \
                         ' or international security and stability.'
    elif cat == '20':
        catdescription = 'To promote and strengthen effective democracies in recipient states and move them along a' \
                         ' continuum toward democratic consolidation.'
    elif cat == '30':
        catdescription = 'To contribute to improvements in the health of people, especially women, children, and' \
                         ' other vulnerable populations in countries of the developing world, through expansion of' \
                         ' basic health services, including family planning; strengthening national health systems,' \
                         ' and addressing global issues and special concerns such as HIV/AIDS and other infectious' \
                         ' diseases.'
    elif cat == '40':
        catdescription = 'Promote equitable, effective, accountable, and sustainable formal and non-formal education' \
                         ' systems and address factors that place individuals at risk for poverty, exclusion,' \
                         ' neglect, or victimization. Help populations manage their risks and gain access to' \
                         ' opportunities that support their full and productive participation in society. Help' \
                         ' populations rebound from temporary adversity, cope with chronic poverty, reduce' \
                         ' vulnerability, and increase self-reliance.'
    elif cat == '50':
        catdescription = 'Generate rapid, sustained, and broad-based economic growth.'
    elif cat == '60':
        catdescription = 'Activities that support the sustainability of a productive and clean environment by:' \
                         ' ensuring that the environment and the natural resources upon which human lives and' \
                         ' livelihoods depend are managed in ways that sustain productivity growth, a healthy' \
                         ' population, as well as the intrinsic spiritual and cultural value of the environment,' \
                         ' and conserving biodiversity and managing natural resources in ways that maintain their' \
                         ' long-term viability and preserve their potential to meet the needs of present and' \
                         ' future generations.'
    elif cat == '70':
        catdescription = 'To save lives, alleviate suffering, and minimize the economic costs of conflict,' \
                         ' disasters and displacement. Humanitarian assistance is provided on the basis of need' \
                         ' according to principles of universality, impartiality and human dignity. It is often' \
                         ' organized by sectors, but requires an integrated, coordinated and/or multi-sectoral' \
                         ' approach to be most effective. Emergency operations will foster the transition from' \
                         ' relief, through recovery, to development, but they cannot and will not replace the' \
                         ' development investments necessary to reduce chronic poverty or establish just' \
                         ' social services.'
    elif cat == '80':
        catdescription = 'To provide the general management support required to ensure completion of U.S. foreign' \
                         ' assistance objectives by facilitating program management, monitoring and evaluation,' \
                         ' and accounting and tracking for costs.'
    else:
        catdescription = 'Unspecified is used only when the category and sector are unknown or cannot be specified.'
    return catdescription


def orgnumber(org):
    """
    Used to return the organization's registered IATI code.
    :param org: The organization to pull the code for
    :return: The organization's IATI code
    """
    if org == 'Department of Agriculture' or org == 'US Department of Agriculture' or org == 'Dept of Agriculture':
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
        orgcode = 'US-GOV-1'
    return orgcode


def sectors_loop(recip, sects, cat):
    """
    Find which activities are the main ones to set as hierarchy 1.
    :param recip: The recipient country or region code.
    :param sects: The sector code in the results file.
    :param cat: The sector category code to ensure that the sectors are within.
    :return sectors: The index numbers corresponding to the rows of the appropriate sectors.
    """
    sectors = list()
    actives = list()
    for i in range(0, len(sects)):
        if "US-GOV-1-" + str(recip[i]) + "-" + str(sects[i])[:2] == cat:
            if str(sects[i])[:2] == cat[-2:]:  # This ensures only groups returned are current category.
                if str(sects[i]) not in actives:
                    actives.append(str(sects[i]))
                    sectors.append([sects[i],i])
                else:
                    for each in sectors:
                        if sects[i] == each[0]:
                            each.append(i)
    return sectors


def results_loop(resultsfile, curcat, iatiactivity):
    """
    Creates the results elements within the XML.
    :param resultsfile: The file to pull the results data from.
    :param curcat: The sector category to pull the sectors for.
    :param iatiactivity: The parent activity to attach the sector elements to.
    :return: None
    """
    sectorgroup = sectors_loop(resultsfile["recipient_code"], resultsfile["sector_code"], curcat)
    print("Sector group: " + str(sectorgroup))
    for resultgroup in sectorgroup:
        print("Result group: " + str(resultgroup))
        # Variables
        hierarchy = '2'
        resLangList = ['en']
        currency = 'USD'
        iden = curcat + "-" + str(resultgroup[0])
        print(iden)
        repOrgR = 'US-GOV-1'
        repOrgT = '10'  # Government
        repText = 'USA'
        rpartOrgRef1 = 'US-GOV-1'
        rpartOrgRole1 = '1'
        rpartOrgText1 = 'USA'
        rpartOrgType1 = '10'  # Government
        rpartOrgRef2 = 'US-GOV-1'  # USAID
        rpartOrgRole2 = '2'
        rpartOrgText2 = 'U.S. Agency for International Development'
        rpartOrgType2 = '10'  # Government
        resTitleText = list()
        resDescText = list()

        # Create the activity group's name
        if str(resultsfile["country"][resultgroup[1]]) != 'nan':
            if str(resultsfile["category"][resultgroup[1]]) != 'nan':
                restitle = 'US-' + str(resultsfile["country"][resultgroup[1]]) + '-' + \
                        rpartOrgText2 + '-' + \
                        str(resultsfile["category"][resultgroup[1]]) + '-' + \
                        str(resultsfile["sector"][resultgroup[1]])
            else:
                restitle = 'US-' + str(resultsfile["country"][resultgroup[1]]) + '-' + \
                        rpartOrgText2 + '-None Found'
        else:
            if str(resultsfile["category"][resultgroup[1]]) != 'nan':
                restitle = 'US-Developing countries, unspecified-' + rpartOrgText2 + '-' + \
                           str(resultsfile["category"][resultgroup[1]]) + '-' + \
                           str(resultsfile["sector"][resultgroup[1]])
            else:
                restitle = 'US-Developing countries, unspecified-' + rpartOrgText2 + '-None Found'
        resTitleText.append(restitle)
        resTitleText.append('')
        resDescText.append(str(resultsfile["program_element_name"][resultgroup[1]]))
        resDescText.append('')

        try:
            ractivityStatusCode = str(int(resultsfile["activity_status"][resultgroup[1]]))
        except ValueError:
            ractivityStatusCode = '1'
        try:
           # Award Transaction Date?
            risodatetime = str(int(resultsfile["result_period_start"][resultgroup[1]]))  # Award Transaction Date
            risodatetimeformat = risodatetime[0:4] + '-' + risodatetime[4:6] + '-' + \
                                 risodatetime[6:8]
            ractivityDateText = ''
        except ValueError:
            risodatetimeformat = str(int(datetime.datetime.utcnow().strftime('%Y')) - 1) + '-10-01'
            ractivityDateText = 'Activity date not available, using Result Indicator Period Start date'
        ractivityDateType = '2'

        # Create the Element Tree for the sector
        active = SubElement(iatiactivity, 'iati-activity', hierarchy=hierarchy,
                            last_h_updated_h_datetime=date,
                            xml__lang=langList[0], default_h_currency=currency)
        identi = SubElement(active, 'iati-identifier')
        identi.text = iden
        report_org = SubElement(active, 'reporting-org', ref=repOrgR,
                                type=repOrgT)
        narrative = SubElement(report_org, 'narrative')
        narrative.text = repText
        resultTitle = SubElement(active, 'title')
        lang_loop(resultTitle, resLangList, resTitleText)
        resultDesc = SubElement(active, 'description')
        lang_loop(resultDesc, resLangList, resDescText)
        resPartOrg1 = SubElement(active, 'participating-org',
                                 ref=rpartOrgRef1, role=rpartOrgRole1,
                                 type=rpartOrgType1)
        narrative = SubElement(resPartOrg1, 'narrative')
        narrative.text = rpartOrgText1
        resPartOrg2 = SubElement(active, 'participating-org',
                                 ref=rpartOrgRef2, role=rpartOrgRole2,
                                 type=rpartOrgType2)
        narrative = SubElement(resPartOrg2, 'narrative')
        narrative.text = rpartOrgText2
        SubElement(active, 'activity-status', code=ractivityStatusCode)
        ractivity_date = SubElement(active, 'activity-date',
                                   iso_h_date=risodatetimeformat,
                                   type=ractivityDateType)
        narrative = SubElement(ractivity_date, 'narrative')
        narrative.text = ractivityDateText

        # Results Loop Here
        resultit = 0
        for result in resultgroup:
            if resultit > 0:
                resulttype = str(int(resultsfile["type_code"][result]))  # 1 = output, 2 = outcome, 3 = impact, 4 = other
                # TODO: Implement when available: resulttext = str(resultsfile["result_title"][result])
                resulttext = str(resultsfile["program_element_name"][result])
                meas = str(int(resultsfile["indicator_measure_code"][result]))  # 1 = unit, 2 = percentage
                indicatortitle = str(resultsfile["indicator_name"][result])
                indstart = str(resultsfile["result_period_start"][result])
                indicatorstart = indstart[0:4] + '-' + indstart[4:6] + '-' + indstart[6:8]
                indend = str(resultsfile["result_period_end"][result])
                indicatorend = indend[0:4]+'-'+indend[4:6]+'-'+indend[6:8]
                try:
                    targetvalue = '{0:.2f}'.format(int(resultsfile["target_result"][result]))
                    actualvalue = '{0:.2f}'.format(int(resultsfile["actual_result"][result]))
                except ValueError:
                    targetvalue = '0.00'
                    actualvalue = '0.00'

                resact = SubElement(active, 'result', type=resulttype)
                rest = SubElement(resact, 'title')
                narrative = SubElement(rest, 'narrative')
                if resulttext != 'nan':
                    narrative.text = resulttext
                indicator = SubElement(resact, 'indicator', measure=meas)
                indtitle = SubElement(indicator, 'title')
                narrative = SubElement(indtitle, 'narrative')
                narrative.text = indicatortitle
                indperiod = SubElement(indicator, 'period')
                SubElement(indperiod, 'period-start', iso_h_date=indicatorstart)
                SubElement(indperiod, 'period-end', iso_h_date=indicatorend)
                SubElement(indperiod, 'target', value=targetvalue)
                SubElement(indperiod, 'actual', value=actualvalue)
            resultit += 1
    return


# Prompt user for filename
filetoopen = input("What is the name of the omb source file? ")
curtime = time.time()
print('Opening OMB file...')
# Read the file
try:
    omb = pandas.read_excel(filetoopen, encoding='utf-8')
except:
    sys.exit("OMB file does not exist.")
# Output the number of rows
print('Total rows: {0}'.format(len(omb)))
# See which headers are available
print(list(omb))

# Prompt user for filename
results = ''
resultsopen = input("What is the name of the results source file? (If no results, then hit enter) ")
if resultsopen != '':
    print('Opening Results file...')
    # Read the file
    try:
        results = pandas.read_excel(resultsopen, encoding='utf-8')
    except:
        sys.exit("Results file does not exist.")
    # Output the number of rows
    print('Total rows: {0}'.format(len(results)))
    # See which headers are available
    print(list(results))

opentime = time.time() - curtime
print('Converting format...')

# Split the recipients into groups here
# This creates the issue of changing the direct file accesses into an index?
# For each in recipientgroup:

# Variable creation
idlist, idawards, isolist = id_loop(omb)
print(idawards)
# This will turn on full dataset dump into one XML. Untab all code after this.
ombActs = activities_loop(idlist)

h1acts = activities_loop(idlist)

# This will turn on the splitting of the file via recipient if you uncomment this and tab everything after these.
ombgrouping = group_split(omb)
for ombActs in ombgrouping:
    print(ombActs)
    # TODO: Change this to 2.02 once we implement the 2.02 changes.
    # http://iatistandard.org/201/upgrades/decimal-upgrade-to-2-02/2-02-changes/
    ver = '2.01'
    fasite = 'http://www.foreignassistance.gov/web/IATI/usg-extension'

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
                partOrgText1 = str(omb["appropriated_agency"][act])
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
                if str(omb["dac_country_name"][act]) != 'nan':
                    try:
                        name = str(omb["dac_country_name"][act])
                        if name == "CÃ´te dâ€™Ivoire":
                            name = "Côte d'Ivoire"
                        elif name == "Lao Peopleâ€™s Democratic Republic":
                            name = "Lao People's Democratic Republic"
                        title = 'US-' + name + '-' +\
                                partOrgText2 + '-' +\
                                str(omb["category_name"][act])
                    except ValueError:
                        name = str(omb["dac_country_name"][act])
                        if name == "CÃ´te dâ€™Ivoire":
                            name = "Côte d'Ivoire"
                        elif name == "Lao Peopleâ€™s Democratic Republic":
                            name = "Lao People's Democratic Republic"
                        title = 'US-' + name + '-' + \
                                partOrgText2 + '-None Found'
                else:
                    try:
                        title = 'US-Worldwide-' + partOrgText2 + '-' +\
                                str(omb["category_name"][act])
                    except ValueError:
                        title = 'US-Worldwide-' + partOrgText2 + '-None Found'
                titleText.append(title)
                titleText.append('')
                # descText.append(str(omb["Implementing Mechanism Purpose Statement"][act]))
                descText.append(catdesc(ident[-2:]))
                descText.append('')

                try:
                    activityStatusCode = str(int(omb["award_status_code"][act]))
                except ValueError:
                    activityStatusCode = '1'
                try:
                    # Award Transaction Date?
                    isodatetime = str(int(omb["award_transaction_date"][act]))
                    isodatetimeformat = isodatetime[0:4] + '-' + isodatetime[4:6] + '-' +\
                        isodatetime[6:8]
                    activityDateText = ''
                except ValueError:
                        isodatetimeformat = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
                        activityDateText = 'No valid date could be found for transaction.'
                activityDateType = '2'
                activityScopeCode = str(int(omb["activity_scope"][act]))

                # Create the Element Tree for the activity grouping
                activity = SubElement(activities, 'iati-activity', hierarchy=hier,
                                      last_h_updated_h_datetime=lastUpdate,
                                      xml__lang=langList[0], default_h_currency=cur)
                identifier = SubElement(activity, 'iati-identifier')
                identifier.text = ident
                reporting_org = SubElement(activity, 'reporting-org', ref=repOrgRef,
                                           type=repOrgType)
                narrative = SubElement(reporting_org, 'narrative')
                narrative.text = repOrgText
                title = SubElement(activity, 'title')
                lang_loop(title, langList, titleText)
                description = SubElement(activity, 'description')
                lang_loop(description, langList, descText)
                participating_org1 = SubElement(activity, 'participating-org',
                                                ref=partOrgRef1, role=partOrgRole1,
                                                type=partOrgType1)
                narrative = SubElement(participating_org1, 'narrative')
                narrative.text = partOrgText1
                participating_org2 = SubElement(activity, 'participating-org',
                                                ref=partOrgRef2, role=partOrgRole2,
                                                type=partOrgType2)
                narrative = SubElement(participating_org2, 'narrative')
                narrative.text = partOrgText2

                SubElement(activity, 'activity-status', code=activityStatusCode)
                activity_date = SubElement(activity, 'activity-date',
                                           iso_h_date=isodatetimeformat,
                                           type=activityDateType)
                narrative = SubElement(activity_date, 'narrative')
                narrative.text = activityDateText
                SubElement(activity, 'activity-scope', code=activityScopeCode)

                relatedList = related_loop(idlist, idawards, ident)

                for rel in relatedList:
                    relType = '2'
                    relRef = idawards[rel]
                    SubElement(activity, 'related-activity', type=relType, ref=relRef)

                if resultsopen != '':
                    results_loop(results, ident, activities)

                # Begin creating the activities
                for relact in relatedList:
                    hier = '3'
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
                    titleText.append(str(omb["implementing_mechanism_title"][relact]))
                    titleText.append('')
                    descText.append(str(omb["iati_narr_lvl_3"][relact]))
                    descText.append('')
                    partOrgText = str(omb["appropriated_agency"][relact])
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
                    partOrgText4 = str(omb["implementing_agent"][relact])
                    partOrgRef4 = orgnumber(partOrgText4)
                    try:
                        partOrgType4 = str(int(omb["implementing_agent_type"][relact]))
                    except ValueError:
                        partOrgType4 = ''
                    try:
                        activityStatusCode = str(int(omb["award_status_code"][relact]))
                    except ValueError:
                        activityStatusCode = '1'

                    # All dates are always "actual". There are no "planned" dates.
                    try:
                        isodatetime = str(int(omb["start_date"][relact]))
                        isodatetimeformatstart = isodatetime[0:4] + '-' +\
                            isodatetime[4:6] + '-' + isodatetime[6:8]
                        activityStartDateText = ''
                    except ValueError:
                        isodatetimeformatstart = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
                        activityStartDateText = str(omb["start_date_narr"][relact])
                    try:
                        isodatetime = str(int(omb["end_date"][relact]))
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
                    activityScopeCode = str(int(omb["activity_scope"][relact]))
                    try:
                        signdate = \
                            str(int(omb["implementing_mech_sign_date"][relact]))
                        signdateformat = signdate[0:4] + '-' + signdate[4:6] + '-' +\
                            signdate[6:8]
                    except ValueError:
                        signdateformat = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'

                    # Put together the first part of the activity element tree
                    activity = SubElement(activities, 'iati-activity', hierarchy=hier,
                                          last_h_updated_h_datetime=lastUpdate,
                                          xml__lang=langList[0], default_h_currency=cur)
                    identifier = SubElement(activity, 'iati-identifier')
                    identifier.text = identity
                    reporting_org = SubElement(activity, 'reporting-org', ref=repOrgRef,
                                               type=repOrgType)
                    narrative = SubElement(reporting_org, 'narrative')
                    narrative.text = repOrgText
                    title = SubElement(activity, 'title')
                    lang_loop(title, langList, titleText)
                    description = SubElement(activity, 'description')
                    lang_loop(description, langList, descText)
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
                    activity_startdate = SubElement(activity, 'activity-date',
                                                    iso_h_date=isodatetimeformatstart,
                                                    type=activityDateTypeStart)
                    if activityStartDateText:
                        narrative = SubElement(activity_startdate, 'narrative')
                        narrative.text = activityStartDateText

                    # All dates are always "actual". There are no "planned" dates.
                    # activity_planend = SubElement(activity, 'activity-date',
                    #                               iso_h_date=isodatetimeformatend,
                    #                               type=activityDateTypePlanEnd)
                    activity_enddate = SubElement(activity, 'activity-date',
                                                  iso_h_date=isodatetimeformatend,
                                                  type=activityDateTypeEnd)
                    if activityEndDateText:
                        narrative = SubElement(activity_enddate, 'narrative')
                        narrative.text = activityStartDateText

                    # Contact information block variables
                    contactType = '1'
                    organisationText = 'U.S. Agency for International Development'
                    personNameText = str(omb["usaid_name"][relact])
                    telephoneText = str(omb["usaid_tel"][relact])
                    emailText = str(omb["usaid_email"][relact])
                    websiteText = str(omb["activity_website"][relact])
                    mailingText = str(omb["usaid_addr"][relact])

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
                    if str(omb["iso_alpha_code"][relact]) == 'nan':
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
                    locsList = location_loop(omb, relact)
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

                    if str(omb["collaboration_type"][relact]) == 'Bilateral':
                        collabCode = '1'
                    else:
                        collabCode = '2'
                    # collabCode = str(int(omb["Collaboration Type Code"][relact]))
                    try:
                        flowType = str(int(omb["flow_type"][relact]))
                    except ValueError:
                        flowType = '0'
                    try:
                        financeType = str(int(omb["finance_type"][relact]))
                    except ValueError:
                        financeType = '0'
                    try:
                        aidType = str(omb["aid_type_code"][relact])
                    except ValueError:
                        aidType = '0'
                    try:
                        tiedCode = str(int(omb["tying_status_of_award"][relact]))
                    except ValueError:
                        tiedCode = '0'
                    # This is for error checking
                    try:
                        if str(int(omb["start_date"][relact])) != 'nan':
                            periodStart = \
                                str(int(omb["start_date"][relact]))
                            periodStartDate = periodStart[0:4] + '-' + periodStart[4:6] + '-' +\
                                periodStart[6:8]
                    except ValueError:
                        periodStartDate = str(int(omb["beginning_fiscal_funding_year"][relact])) + '-10-01'
                    try:
                        if str(int(omb["end_date"][relact])) != 'nan':
                            periodEnd = \
                                str(int(omb["end_date"][relact]))
                            periodEndDate = periodEnd[0:4] + '-' + periodEnd[4:6] + '-' +\
                                periodEnd[6:8]
                    except ValueError:
                        try:
                            if str(int(omb["ending_fiscal_funding_year"][relact])) != 'nan':
                                periodEndDate = str(int(omb["ending_fiscal_funding_year"][relact])) + '-09-30'
                        except ValueError:
                            periodEndDate = ''
                    budgetValueDate = periodStartDate
                    try:
                        if str(int(omb["total_allocations"][relact])) != 'nan':
                            budgetAmount = '{0:.2f}'.format(omb["total_allocations"][relact])
                        else:
                            budgetAmount = '0.00'
                    except ValueError:
                        budgetAmount = '0.00'

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
                        budget = SubElement(activity, 'budget')
                        SubElement(budget, 'period-start', iso_h_date=periodStartDate)
                        SubElement(budget, 'period-end', iso_h_date=periodEndDate)
                        budgetValue = SubElement(budget, 'value', currency=cur,
                                                 value_h_date=budgetValueDate)
                        budgetValue.text = budgetAmount

                    # Create the list of transactions for a specific activity
                    transList = trans_loop(idawards, idawards[relact])

                    # These two make it easier to determine
                    # if a 0 has been put in for com/dis
                    # If changed to True, then a 0 transaction will not be put into XML
                    comMarker = True
                    disMarker = True

                    # Loop through the transactions related to the activity
                    for trans in transList:
                        # Variables that depend on entries
                        # If the disbursement has a value, set value to disbursement.
                        try:
                            transAmount = omb["award_transaction_value"][trans]
                            # transAmount = int(omb["Obligation Amount"][trans])
                            # transAmount = int(omb["Disbursement Amount"][trans])
                            valueAmount = '{0:.2f}'.format(transAmount)
                        except ValueError:
                            valueAmount = '0.00'
                        transType = str(omb["award_transaction_type"][trans])
                        if transType == "Commitment":
                            transaction_code = '2'
                        elif transType == "Disbursement":
                            transaction_code = '3'
                        else:
                            transaction_code = '0'
                        try:
                            valuedate = str(int(omb["award_transaction_date"][trans]))
                            value_datetime = valuedate[0:4] + '-' + valuedate[4:6] + \
                                '-' + valuedate[6:8]
                        except ValueError:
                            value_datetime = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
                        regAccCode = str(int(omb["treasury_reg_account_code"][trans]))
                        mainAccCode = str(int(omb["treasury_main_account_code"][trans]))
                        mainText = str(omb["treasury_main_account_title"][trans])
                        try:
                            fundingYearBegin = \
                                str(int(omb["beginning_fiscal_funding_year"][trans]))
                        except ValueError:
                            fundingYearBegin = str(int(datetime.datetime.utcnow().strftime('%Y'))-1)
                        try:
                            fundingYearEnd = \
                                str(int(omb["ending_fiscal_funding_year"][trans]))
                        except ValueError:
                            fundingYearEnd = str(datetime.datetime.utcnow().strftime('%Y'))

                        # Make sure there is exactly 1 transaction value of 0.00 for Com if needed
                        if transaction_code == '2':
                            if (comMarker is True and valueAmount != '0.00') or (comMarker is False and valueAmount == '0.00'):
                                # Set the elements
                                transaction = SubElement(activity, 'transaction')
                                transaction_type = SubElement(transaction, 'transaction-type',
                                                              code=transaction_code)
                                transaction_date = SubElement(transaction, 'transaction-date',
                                                              iso_h_date=value_datetime)
                                value = SubElement(transaction, 'value',
                                                   value_h_date=value_datetime)
                                value.text = valueAmount

                                try:
                                    disbChan = str(int(omb["disbursement_channel"][trans]))
                                except ValueError:
                                    disbChan = '0'

                                # Sectors
                                try:
                                    dacCode = str(int(omb["dac_purpose_code"][trans]))
                                except ValueError:
                                    dacCode = '0'
                                try:
                                    sectorCode = str(int(omb["us_sector_code"][trans]))
                                except ValueError:
                                    sectorCode = '0'
                                dacVocab = '1'
                                dacText = str(omb["dac_purpose_name"][trans])
                                sectorVocab = '99'
                                sectorText = str(omb["aid_type_name"][trans])

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
                            comMarker = True

                        # Make sure there is exactly 1 transaction value of 0.00 for Disb if needed
                        if transaction_code == '3':
                            if (disMarker is True and valueAmount != 0) or (disMarker is False and valueAmount == 0):
                                # Set the elements
                                transaction = SubElement(activity, 'transaction')
                                transaction_type = SubElement(transaction, 'transaction-type',
                                                              code=transaction_code)
                                transaction_date = SubElement(transaction, 'transaction-date',
                                                              iso_h_date=value_datetime)
                                value = SubElement(transaction, 'value',
                                                   value_h_date=value_datetime)
                                value.text = valueAmount

                                try:
                                    disbChan = str(int(omb["disbursement_channel"][trans]))
                                except ValueError:
                                    disbChan = '0'

                                # Sectors
                                try:
                                    dacCode = str(int(omb["dac_purpose_code"][trans]))
                                except ValueError:
                                    dacCode = '0'
                                try:
                                    sectorCode = str(int(omb["us_sector_code"][trans]))
                                except ValueError:
                                    sectorCode = '0'
                                dacVocab = '1'
                                dacText = str(omb["dac_purpose_name"][trans])
                                sectorVocab = '99'
                                sectorText = str(omb["aid_type_name"][trans])

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
                    docsList = docs_loop(omb, relact)
                    for doc in docsList:
                        # FIX: url and format get flipped somehow?
                        document = SubElement(activity, 'document-link', format=doc[2], url=doc[1])
                        title = SubElement(document, 'title')
                        narrative = SubElement(title, 'narrative')
                        narrative.text = doc[0]
                        category = SubElement(document, 'category', code=doc[3])
                        lang = SubElement(document, 'language', code=doc[4])

                    # TODO: Insert code for Contract Links here
                    # document-link code=A11
                    # Concatenate "https://www.usaspending.gov/Pages/AdvancedSearch.aspx?k=" + field
                    # if str(omb["stripped award"][relact]) != nan:
                    #  contractlink = (link) + str(omb["stripped award"][relact])
                    #  contract = subelement(activity, 'document-link', format=html, url=contractlink)
                    #  subelement(contract, 'category', code="A11")
                    #  subelement(contract, 'language', code="en")

                    relType = '1'
                    relRef = ident
                    SubElement(activity, 'related-activity', type=relType, ref=relRef)
                    # This assumes that there will never be any conditions.
                    # This is currently the case, however, this may eventually change.
                    conditions = SubElement(activity, 'conditions', attached="0")
                    SubElement(activity, 'usg__mechanism-signing-date',
                               iso_h_date=signdateformat)

                    # Extra fields requested by State
                    try:
                        duns = str(int(omb["implementing_agent_duns_number"][relact]))
                    except ValueError:
                        duns = 'nan'
                    tec = '{0:.2f}'.format(omb["tec1"][relact])
                    stateloc = str(omb["state_location"][relact])
                    if stateloc == "CÃ´te d'Ivoire":
                        stateloc = "Côte d'Ivoire"
                    elif stateloc == "Lao Peopleâ€™s Democratic Republic":
                        stateloc = "Lao People's Democratic Republic"

                    if (duns != 'nan'):
                        dunselement = SubElement(activity, 'usg__duns-number')
                        narrative = SubElement(dunselement, 'narrative')
                        narrative.text = duns
                    if (tec != 'nan'):
                        tecelement = SubElement(activity, 'usg__tec1')
                        narrative = SubElement(tecelement, 'narrative')
                        narrative.text = tec
                    if (stateloc != 'nan'):
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
    output_file = open('export/' + time.strftime("%m-%d-%Y") + '/iati-activities-'+str(omb["state_location"][int(ombActs[1])])+'.xml', 'w', encoding='utf-8')
    # This line is for country codes
    # output_file = open('export/' + time.strftime("%m-%d-%Y") + '/iati-activities-'+ombActs[0]+'.xml', 'w', encoding='utf-8')
    output_file.write(prettify(activities).replace("__", ":").replace("_h_", "-"))
    output_file.close()
    finaltime = time.time() - curtime
    print('Opening Time: ' + str(opentime))
    print('Convert Time: ' + str(converttime - opentime))
    print('Write Time: ' + str(finaltime - converttime))
    print('Run time: ' + str(finaltime))
    print('Average time per main activity: ' +
          str((converttime - opentime)/len(ombActs)))
print('Zipping...')
shutil.make_archive('export-' + time.strftime("%m-%d-%Y"), 'zip', 'export/' + time.strftime("%m-%d-%Y") + '/')
print('Complete!')