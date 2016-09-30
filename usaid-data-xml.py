import datetime
import time
import sys
from xml.etree import ElementTree
from xml.etree.ElementTree import Element, SubElement
from xml.dom import minidom
import pandas

__author__ = "Timothy Cameron"
__email__ = "tcameron@devtechsys.com"
__date__ = "5-26-2016"
__version__ = "0.8"

# TODO: Use the ISO codes to determine languages to translate into.
# FIX: Create method to create the fields, as opposed to them being
# populated in the main method
# TODO: Make sure all fields will line up with the max number points.


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
    # USAID is identified as US-1 within PWYF.
    code = '1'
    for i in range(0, len(ombfile.index)):
        try:
            country = str(int(ombfile["DAC Regional Code"][i]))
        except ValueError:
            if str(ombfile["DAC Country Code"][i]) != '998':
                try:
                    country = str(ombfile["ISO Alpha Code"][i])
                except ValueError:
                    country = 'XXX'
            else:
                country = '998'
        try:
            agencytype = str(int(omb["U.S. Government Sector Code"][i]))[0:2]
        except ValueError:
            agencytype = '10'
        awardid = str(omb["Implementing Mechanism ID"][i])
        entry = 'US' + '-' + code + '-' + country + '-' + agencytype
        ids.append(entry)
        entry += '-' + awardid
        idswawards.append(entry)
        isos.append(country)
    return ids, idswawards, isos


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


# TODO: Shorten this
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
    print('in docs_loop')
    if str(ombfile["Evaluation Title 1"][award]) != 'nan':
        print("Entered Eval Title 1")
        tempdoc.append(str(ombfile["Evaluation Title 1"][award]))
        tempdoc.append(str(ombfile["Evaluation Link 1"][award]))
        tempdoc.append(str(ombfile["Evaluation File Format"][award]))
        tempdoc.append(str(ombfile["Evaluation Document Category"][award]))
        tempdoc.append(str(ombfile["Evaluation Language 1"][award]))
        docs.append(tempdoc)
        tempdoc = []
    if str(ombfile["Evaluation Title 2"][award]) != 'nan':
        print("Entered Eval Title 2")
        tempdoc.append(str(ombfile["Evaluation Title 2"][award]))
        tempdoc.append(str(ombfile["Evaluation Link 2"][award]))
        tempdoc.append(str(ombfile["Evaluation File Format"][award]))
        tempdoc.append(str(ombfile["Evaluation Document Category"][award]))
        tempdoc.append(str(ombfile["Evaluation Language 2"][award]))
        docs.append(tempdoc)
        tempdoc = []
    if str(ombfile["Evaluation Title 3"][award]) != 'nan':
        print("Entered Eval Title 3")
        tempdoc.append(str(ombfile["Evaluation Title 3"][award]))
        tempdoc.append(str(ombfile["Evaluation Link 3"][award]))
        tempdoc.append(str(ombfile["Evaluation File Format"][award]))
        tempdoc.append(str(ombfile["Evaluation Document Category"][award]))
        tempdoc.append(str(ombfile["Evaluation Language 3"][award]))
        docs.append(tempdoc)
        tempdoc = []
    if str(ombfile["Evaluation Title 4"][award]) != 'nan':
        print("Entered Eval Title 4")
        tempdoc.append(str(ombfile["Evaluation Title 4"][award]))
        tempdoc.append(str(ombfile["Evaluation Link 4"][award]))
        tempdoc.append(str(ombfile["Evaluation File Format"][award]))
        tempdoc.append(str(ombfile["Evaluation Document Category"][award]))
        tempdoc.append(str(ombfile["Evaluation Language 4"][award]))
        docs.append(tempdoc)
        tempdoc = []
    if str(ombfile["Impact Appraisal Title 1"][award]) != 'nan':
        print("Entered Impact Appraisal Title 1")
        tempdoc.append(str(ombfile["Impact Appraisal Title 1"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Link 1"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal File Format"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Document Category"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Language 1"][award]))
        docs.append(tempdoc)
        tempdoc = []
    if str(ombfile["Impact Appraisal Title 2"][award]) != 'nan':
        print("Entered Impact Appraisal Title 2")
        tempdoc.append(str(ombfile["Impact Appraisal Title 2"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Link 2"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal File Format"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Document Category"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Language 2"][award]))
        docs.append(tempdoc)
        tempdoc = []
    if str(ombfile["Impact Appraisal Title 3"][award]) != 'nan':
        print("Entered Impact Appraisal Title 3")
        tempdoc.append(str(ombfile["Impact Appraisal Title 3"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Link 3"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal File Format"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Document Category"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Language 3"][award]))
        docs.append(tempdoc)
        tempdoc = []
    if str(ombfile["Impact Appraisal Title 4"][award]) != 'nan':
        print("Entered Impact Appraisal Title 4")
        tempdoc.append(str(ombfile["Impact Appraisal Title 4"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Link 4"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal File Format"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Document Category"][award]))
        tempdoc.append(str(ombfile["Impact Appraisal Language 4"][award]))
        docs.append(tempdoc)
    print('award ' + str(award))
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


# Prompt user for filename
filetoopen = input("What is the name of the source file? ")
curtime = time.time()
print('Opening OMB file...')
# Read the file
try:
    omb = pandas.read_excel(filetoopen, encoding='utf-8')
except:
    sys.exit("File does not exist.")
# Output the number of rows
print('Total rows: {0}'.format(len(omb)))
# See which headers are available
print(list(omb))

opentime = time.time() - curtime
print('Converting format...')

# Variable creation
idlist, idawards, isolist = id_loop(omb)
ombActs = activities_loop(idlist)
date = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3]+'Z'
# TODO: Change this to 2.02 once we implement the 2.02 changes. (Humanitarian)
ver = '2.01'
fasite = 'http://www.foreignassistance.gov/web/IATI/usg-extension'

activities = Element('iati-activities', version=ver,
                     generated_h_datetime=date, xmlns__usg=fasite)

# Start creating the hierarchy 1 groupings
l = 1
print(len(ombActs))
# For loop for the amount of activities
for act in ombActs:
    print(l)
    l += 1
    # Variables
    hier = '1'
    lastUpdate = date
    # TODO: Presumably, pull in data from alternate translation fields?
    langList = ['en']
    cur = 'USD'
    ident = idlist[act]
    repOrgRef = 'US-USAGOV'
    repOrgType = '10'  # Government
    repOrgText = 'USA'
    partOrgRef1 = 'US-USAGOV'
    partOrgRole1 = '1'
    partOrgText1 = 'USA'
    partOrgType1 = '10'  # Government
    partOrgRef2 = 'US-1'  # USAID
    partOrgRole2 = '2'
    partOrgText2 = 'U.S. Agency for International Development'
    partOrgType2 = '10'  # Government
    titleText = list()
    descText = list()

    # Create the activity group's name
    if str(omb["DAC Country Name"][act]) != 'nan':
        try:
            title = 'US-' + str(omb["DAC Country Name"][act]) + '-' +\
                    partOrgText2 + '-' +\
                    str(int(omb["U.S. Government Category Name"][act]))
        except ValueError:
            title = 'US-' + str(omb["DAC Country Name"][act]) + '-' + \
                    partOrgText2 + '-None Found'
    else:
        try:
            title = 'US-Worldwide-' + partOrgText2 + '-' +\
                    str(int(omb["U.S. Government Category Name"][act]))
        except ValueError:
            title = 'US-Worldwide-' + partOrgText2 + '-None Found'
    titleText.append(title)
    titleText.append('')
    # descText.append(str(omb["Implementing Mechanism Purpose Statement"][act]))
    # TODO: Get a description of each different category
    descText.append('This is a place holder category description.')
    descText.append('')

    try:
        activityStatusCode = str(int(omb["Reporting Status"][act]))
    except ValueError:
        activityStatusCode = '1'
    try:
        # Award Transaction Date?
        isodatetime = str(int(omb["Award Transaction Date"][act]))
        isodatetimeformat = isodatetime[0:4] + '-' + isodatetime[4:6] + '-' +\
            isodatetime[6:8]
        activityDateText = ''
    except ValueError:
            isodatetime = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
            activityDateText = 'No valid date could be found for transaction.'
    activityDateType = '2'
    activityScopeCode = '4'

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

    # Begin creating the activities
    for relact in relatedList:
        hier = '3'
        lastUpdate = date
        langList = ['en']
        cur = 'USD'
        countryinit = isolist[relact]
        identity = idawards[relact]
        repOrgRef = 'US-USAGOV'
        repOrgType = '10'  # Government
        repOrgText = 'USA'
        # These will need to be looped through and placed into narratives.
        titleText = list()
        descText = list()
        titleText.append(str(omb["Implementing Mechanism Title"][relact]))
        titleText.append('')
        descText.append(str(omb["Implementing Mechanism Purpose Statement"][relact]))
        descText.append('')
        partOrgRef = 'US-USAGOV'
        partOrgRole = '1'
        partOrgText = 'USA'
        partOrgRef2 = 'US-1'
        partOrgRole2 = '2'
        partOrgText2 = 'U.S. Agency for International Development'
        partOrgType2 = '10'
        partOrgRole3 = '3'
        partOrgText3 = str(omb["Appropriated Agency"][act])
        # These may change, depending on input.
        # TODO: If more organizations become used, they will need impl.
        if partOrgText3 == "Dept of State":
            partOrgRef3 = 'US-11'
        elif partOrgText3 == "Dept of Agriculture":
            partOrgRef3 = 'US-2'
        elif partOrgText3 == "MCC":
            partOrgRef3 = 'US-18'
        else:
            partOrgRef3 = 'US-1'  # USAID
        partOrgType3 = '10'
        partOrgRole4 = '4'
        partOrgText4 = str(omb["Implementing Agent"][act])
        if partOrgText4 == "US Department of State":
            partOrgRef4 = 'US-11'
        elif partOrgText4 == "US Department of Treasury":
            partOrgRef4 = 'US-6'
        elif partOrgText4 == "US Department of Agriculture":
            partOrgRef4 = 'US-2'
        elif partOrgText4 == "USAID Mission" or partOrgText4 == "USAID/Washington":
            partOrgRef4 = 'US-1'
        else:
            partOrgRef4 = ''
        try:
            partOrgType4 = str(int(omb["Implementing Agent Type"][act]))
        except ValueError:
            partOrgType4 = ''
        try:
            activityStatusCode = str(int(omb["Reporting Status"][relact]))
        except ValueError:
            activityStatusCode = '1'
        # TODO: Check for dates to determine between planned and actual
        # All dates are always "actual". There are no "planned" dates.
        try:
            isodatetime = str(int(omb["Start Date"][act]))
            isodatetimeformatstart = isodatetime[0:4] + '-' +\
                isodatetime[4:6] + '-' + isodatetime[6:8]
            activityStartDateText = ''
        except ValueError:
            isodatetimeformatstart = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
            activityStartDateText = 'No valid date could be found for' \
                                    ' activity.'
        try:
            isodatetime = str(int(omb["End Date"][act]))
            isodatetimeformatend = isodatetime[0:4] + '-' + isodatetime[4:6] +\
                '-' + isodatetime[6:8]
            activityEndDateText = ''
        except ValueError:
            isodatetimeformatend = datetime.datetime.utcnow().strftime('%Y') + '-10-01'
            activityEndDateText = 'No valid date could be found for' \
                                  ' activity.'
        activityDateTypePlanStart = '1'
        activityDateTypeStart = '2'
        activityDateTypePlanEnd = '3'
        activityDateTypeEnd = '4'
        activityScopeCode = '4'
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
        # TODO: Implement planned start date
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
        # TODO: Implement planned end date
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
        if str(omb["DAC Regional Code"][relact]) != 'nan':
            recipientRegionCode = countryinit
            recipient_region = SubElement(activity, 'recipient-region',
                                          percentage=recipientCountryPercentage,
                                          code=recipientRegionCode)
        else:
            recipientCountryCode = countryinit
            recipient_country = SubElement(activity, 'recipient-country',
                                           percentage=recipientCountryPercentage,
                                           code=recipientCountryCode)
        collabCode = str(int(omb["Collaboration Type Code"][relact]))
        try:
            flowType = str(int(omb["Flow Type"][relact]))
        except ValueError:
            flowType = 0
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
        periodStartDate = str(int(datetime.datetime.utcnow().strftime('%Y'))-1) + '-10-01'
        periodEndDate = str(datetime.datetime.utcnow().strftime('%Y')) + '-09-30'
        budgetValueDate = str(datetime.datetime.utcnow().strftime('%Y')) + '-09-30'
        try:
            budgetAmount = '{0:.2f}'.format(int(omb["Total allocations"][relact]))
        except ValueError:
            budgetAmount = '0.00'

        # Create the pre-transaction types element tree
        collaboration_type = SubElement(activity, 'collaboration-type',
                                        code=collabCode)
        if flowType != 0:
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
        # FIX: Grab relevant info from the new sources and set vals.
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
                transAmount = int(omb["Award Transaction Value"][trans])
                # transAmount = int(omb["Obligation Amount"][trans])
                # transAmount = int(omb["Disbursement Amount"][trans])
                valueAmount = '{0:.2f}'.format(transAmount)
            except ValueError:
                valueAmount = '0.00'
            transType = str(omb["Award Transaction Type"][trans])
            if transType == "Commitment":
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
                        disbChan = str(int(omb["Disbursement Channel"][relact]))
                    except ValueError:
                        disbChan = '0'

                    # Sectors
                    try:
                        dacCode = str(int(omb["DAC Purpose Code"][relact]))
                    except ValueError:
                        dacCode = '0'
                    try:
                        sectorCode = str(int(omb["U.S. Government Sector Code"][relact]))
                    except ValueError:
                        sectorCode = '0'
                    dacVocab = '1'
                    # TODO: Once files are received with this information, change to reflect appr. info
                    dacText = str(omb["DAC Purpose Name"][relact])
                    sectorVocab = '99'
                    sectorText = str(omb["Aid Type Name"][relact])

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
                        disbChan = str(int(omb["Disbursement Channel"][relact]))
                    except ValueError:
                        disbChan = '0'

                    # Sectors
                    try:
                        dacCode = str(int(omb["DAC Purpose Code"][relact]))
                    except ValueError:
                        dacCode = '0'
                    try:
                        sectorCode = str(int(omb["U.S. Government Sector Code"][relact]))
                    except ValueError:
                        sectorCode = '0'
                    dacVocab = '1'
                    # TODO: Once files are received with this information, change to reflect appr. info
                    dacText = str(omb["DAC Purpose Name"][relact])
                    sectorVocab = '99'
                    sectorText = str(omb["Aid Type Name"][relact])

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
        print(docsList)
        for doc in docsList:
            # FIX: url and format get flipped somehow?
            document = SubElement(activity, 'document-link', format=doc[2], url=doc[1])
            title = SubElement(document, 'title')
            narrative = SubElement(title, 'narrative')
            narrative.text = doc[0]
            category = SubElement(document, 'category', code=doc[3])
            lang = SubElement(document, 'language', code=doc[4])

        relType = '1'
        relRef = ident
        SubElement(activity, 'related-activity', type=relType, ref=relRef)
        # This assumes that there will never be any conditions.
        # This is currently the case, however, this may eventually change.
        conditions = SubElement(activity, 'conditions', attached="0")
        SubElement(activity, 'usg__mechanism-signing-date',
                   iso_h_date=signdateformat)

# End of run processing and time keeping stats.
converttime = time.time() - curtime
print('Writing file...')

output_file = open('iati-activities-output.xml', 'w', encoding='utf-8')
output_file.write(prettify(activities).replace("__", ":").replace("_h_", "-"))
output_file.close()
finaltime = time.time() - curtime
print('Opening Time: ' + str(opentime))
print('Convert Time: ' + str(converttime - opentime))
print('Write Time: ' + str(finaltime - converttime))
print('Run time: ' + str(finaltime))
print('Average time per main activity: ' +
      str((converttime - opentime)/len(ombActs)))
print('Complete!')
