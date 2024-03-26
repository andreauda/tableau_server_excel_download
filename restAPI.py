import requests
import xml.etree.ElementTree as ET
from urllib3.exceptions import InsecureRequestWarning

import tableau_details as td


requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)

# The namespace for the REST API is 'http://tableausoftware.com/api' for TBS 9.0
# or 'http://tableau.com/api' for TBS 9.1 or later
XMLNS = {'t' : 'http://tableau.com/api'}
VERSION = td.VERSION

class ApiCallError(Exception):
    pass
    
    
class UserDefinedFieldError(Exception):
	""" UserDefinedFieldError """
	pass

def sign_in(server, username, password, site=""):
    """
    Signs in to the server specified with the given credentials
    'server'   specified server address
    'username' is the name (not ID) of the user to sign in as.
               Note that most of the functions in this example require that the user
               have server administrator permissions.
    'password' is the password for the user.
    'site'     is the ID (as a string) of the site on the server to sign in to. The
               default is "", which signs in to the default site.
    Returns the authentication token and the site ID.
    """
    url = server + "/api/{0}/auth/signin".format(VERSION)

    # Builds the request
    xml_request = ET.Element('tsRequest')
    credentials_element = ET.SubElement(xml_request, 'credentials', name=username, password=password)
    ET.SubElement(credentials_element, 'site', contentUrl=site)
    xml_request = ET.tostring(xml_request)

    # Make the request to server
    server_response = requests.post(url, data=xml_request)
    _check_status(server_response, 200)

    # ASCII encode server response to enable displaying to console
    server_response = _encode_for_display(server_response.text)

    # Reads and parses the response
    parsed_response = ET.fromstring(server_response)

    # Gets the auth token and site ID
    token = parsed_response.find('t:credentials', namespaces=XMLNS).get('token')
    site_id = parsed_response.find('.//t:site', namespaces=XMLNS).get('id')
    user_id = parsed_response.find('.//t:user', namespaces=XMLNS).get('id')
    
    return token, site_id, user_id


def sign_out(server, auth_token):
    """
    Destroys the active session and invalidates authentication token.
    'server'        specified server address
    'auth_token'    authentication token that grants user access to API calls
    """
    url = server + "/api/{0}/auth/signout".format(VERSION)
    server_response = requests.post(url, headers={'x-tableau-auth': auth_token}, verify=False)
    _check_status(server_response, 204)

    return
	
	
# Method to get the id for a given workbook name
def get_wb_id(server, auth_token, site_id, wb_name):
    """
    Returns the workbook id for the workbook name
    """
    url = server + "/api/{0}/sites/{1}/workbooks".format(VERSION, site_id)
    server_response = requests.get(url, headers={'x-tableau-auth': auth_token})
    _check_status(server_response, 200)
    xml_response = ET.fromstring(_encode_for_display(server_response.text))

    wbs = xml_response.findall('.//t:workbook', namespaces=XMLNS)
    
    for wb in wbs:
        #print(wb.get('name'))
        if wb.get('name') == wb_name:
            return wb.get('id')
    error = "Workbook named '{0}' not found.".format(wb_name)
    error = "Workbook named " + str(wb_name) + " not found."
    raise LookupError(error)
    
    
# Method to get the id for a given view name
def get_view_id(server, auth_token, site_id, view_name, pagesize, pagenum): 
    """
    Returns the view id for the view name
    """
    url = server + "/api/{0}/sites/{1}/views?pageSize={2}&pageNumber={3}".format(VERSION, site_id, pagesize, pagenum)
    server_response = requests.get(url, headers={'x-tableau-auth': auth_token})
    _check_status(server_response, 200)
    xml_response = ET.fromstring(_encode_for_display(server_response.text))

    views = xml_response.findall('.//t:view', namespaces=XMLNS)
    
    for view in views:
        #print(wb.get('name'))
        if view.get('name') == view_name:
            return view.get('id')
    error = "View named '{0}' not found.".format(view_name)
    raise LookupError(error)
    
    
# Method to download the given view in excel format
def download_excel_view(server, auth_token, site_id, view_id):
    """
    Downloads the view in excel format
        max-age-minutes	(Optional) The maximum number of minutes an .xlsx file will be cached on the server before being refreshed. 
        To prevent multiple .xlsx requests from overloading the server, the shortest interval you can set is one minute. 
        There is no maximum value, but the server job enacting the caching action may expire before a long cache period is reached.
    """
    url = server + "/api/{0}/sites/{1}/views/{2}/crosstab/excel?maxAge=10".format(VERSION, site_id, view_id) #10minuti massimo di cache, dopo refresh(?)
    server_response = requests.get(url, headers={'x-tableau-auth': auth_token})
    _check_status(server_response, 200)

    return server_response
	
	
def _encode_for_display(text):
    """
    Encodes strings so they can display as ASCII in a Windows terminal window.
    This function also encodes strings for processing by xml.etree.ElementTree functions.
    Returns an ASCII-encoded version of the text.
    Unicode characters are converted to ASCII placeholders (for example, "?").
    """
    
    return text.encode('ascii', errors="backslashreplace").decode('utf-8')


def _check_status(server_response, success_code):
    """
    Checks the server response for possible errors.
    'server_response'       the response received from the server
    'success_code'          the expected success code for the response
    Throws an ApiCallError exception if the API call fails.
    """
    if server_response.status_code != success_code:
        parsed_response = ET.fromstring(server_response.text)

        # Obtain the 3 xml tags from the response: error, summary, and detail tags
        error_element = parsed_response.find('t:error', namespaces=XMLNS)
        summary_element = parsed_response.find('.//t:summary', namespaces=XMLNS)
        detail_element = parsed_response.find('.//t:detail', namespaces=XMLNS)

        # Retrieve the error code, summary, and detail if the response contains them
        code = error_element.get('code', 'unknown') if error_element is not None else 'unknown code'
        summary = summary_element.text if summary_element is not None else 'unknown summary'
        detail = detail_element.text if detail_element is not None else 'unknown detail'
        error_message = '{0}: {1} - {2}'.format(code, summary, detail)
        raise ApiCallError(error_message)
        
    return