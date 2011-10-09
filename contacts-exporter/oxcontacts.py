#!/usr/bin/python
# vim: set expandtab ts=4:

# Copyright (c) 2011, Rooslan S. Khayrov <khayrov@gmail.com>
#
# Permission to use, copy, modify, and/or distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
# 
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.

import sys
from optparse import OptionParser
import time
from getpass import getpass
from urlparse import urlparse
import httplib
from base64 import b64encode
from xml.dom.minidom import parseString


ox_scheme = ''
ox_host = ''
ox_basepath = ''
user = ''
password = ''
output_filename = None


def basic_auth():
    return 'Basic ' + b64encode(user + ':' + password)


propfind_xml_pattern = '''<?xml version="1.0" encoding="utf-8"?>
<D:propfind xmlns:D="DAV:">
    <D:prop xmlns:ox="http://www.open-xchange.org">
        <ox:objectmode>MODIFIED</ox:objectmode>
        <ox:lastsync>0</ox:lastsync>
        %s
    </D:prop>
</D:propfind>'''


def element_text(element):
    content = element.firstChild
    if content and content.nodeType == content.TEXT_NODE:
        return content.data
    else:
        return ''


OX_XMLNS = 'http://www.open-xchange.org'


class OXError(Exception):
    pass


class OXConnection(object):

    def __init__(self):
        if ox_scheme == 'http':
            self.conn = httplib.HTTPConnection(ox_host)
        elif ox_scheme == 'https':
            self.conn = httplib.HTTPSConnection(ox_host)
        else:
            assert False, 'Unknown scheme'

        self.headers = {'Authorization' : basic_auth(),
            'Content-type' : 'application/xml; charset="utf-8"' }

    def close(self):
        self.conn.close()

    def __del__(self):
        self.close()

    def propfind(self, url, xml):
        if url[:1] != '/':
            url = '/' + url

        self.conn.request('PROPFIND', ox_basepath + url, xml, self.headers)
        rsp = self.conn.getresponse()

        if rsp.status != 207:
            msg = '207 Multi-Status expected, got %d %s. Not a WebDAV URL?' % \
                (rsp.status, rsp.reason)
            raise OXError, msg

        return rsp.read()

    def list_contact_folders(self):
        result = []

        xml = self.propfind('/servlet/webdav.folders', propfind_xml_pattern % '')
        dom = parseString(xml)

        for prop_element in dom.getElementsByTagNameNS('DAV:', 'prop'):
            module_element = prop_element.getElementsByTagNameNS(OX_XMLNS, 'module')[0]
            oid_element = prop_element.getElementsByTagNameNS(OX_XMLNS, 'object_id')[0]

            if element_text(module_element) == 'contact':
                result.append(element_text(oid_element))

        return result

    def get_contact_folder_contents(self, folder_id):
        result = []

        xml = self.propfind('/servlet/webdav.contacts',
            propfind_xml_pattern % ('<ox:folder_id>%s</ox:folder_id>' % folder_id))
        dom = parseString(xml)

        for prop_element in dom.getElementsByTagNameNS('DAV:', 'prop'):
            props = {}

            for field_element in prop_element.childNodes:
                if field_element.nodeType == field_element.ELEMENT_NODE and \
                        field_element.namespaceURI == OX_XMLNS:
                    props[field_element.localName] = element_text(field_element)

            result.append(props)

        return result


def vcard_bday(props):
    ts = props.get('birthday')
    if ts:
        tm = time.gmtime(float(ts) / 1000)
        return time.strftime('%Y%m%d', tm)
    else:
        return None

def vcard_ref(props):
    ts = props.get('last_modified')
    if ts:
        tm = time.gmtime(float(ts) / 1000)
        return time.strftime('%Y%m%dT%H%M%SZ', tm)
    else:
        return None

vcard_props_mapping = {
    'FN' : 'displayname',
    'N' : lambda(props): '%s;%s;%s;%s;%s' % (props['last_name'].strip(),
        props.get('first_name', '').strip(), props.get('second_name', '').strip(),
        props.get('title', ''), props.get('suffix', '')),
    'EMAIL;TYPE="work"' : 'email1',
    'ORG' : lambda(props) : '%s;%s' % (
        props.get('company', ''), props.get('department', '')),
    'ROLE' : 'position',
    'TEL;TYPE="voice,work"' : 'phone_business',
    'TEL;type="fax,work"' : 'fax_business',
    'TEL;type="voice,cell"' : 'mobile1',
    'TEL;type="voice,home"' : 'phone_home',
    'ADR;type="work"' : lambda(props): '%s;%s;%s;%s;%s;%s;%s' % (
        '', '', props.get('business_street', ''),
        props.get('business_city', ''), props.get('business_state', ''),
        props.get('business_postal_code', ''),
        props.get('business_country', '')),
    'X-EVOLUTION-MANAGER' : 'managers_name',
    'BDAY' : vcard_bday,
    'REF' : vcard_ref
}


def make_vcard(props):
    result = 'BEGIN:VCARD\r\nVERSION:4.0\r\n'
    for field, mapping in vcard_props_mapping.items():
        value = None
        if callable(mapping):
            value = mapping(props)
        else:
            value = props.get(mapping, '')
        if value and value.replace(';', ''):
            result += '%s:%s\r\n' % (field, value)
    result += 'END:VCARD\r\n'
    return result


def parse_ox_url(url):
    global ox_scheme
    global ox_host
    global ox_basepath

    parsed_url = urlparse(url)
    if not parsed_url.scheme:
        parsed_url= urlparse('https://' + url)

    if parsed_url.scheme not in ('http', 'https'):
        return False
    ox_scheme = parsed_url.scheme

    ox_host = parsed_url.netloc
    ox_basepath = parsed_url.path

    return True


def init_options():
    global user
    global password
    global ox_scheme
    global ox_host
    global ox_basepath
    global output_filename

    usage = 'Usage: %prog [options] <OpenXchange URL>'
    parser = OptionParser(usage=usage)
    parser.add_option('-u', '--user', metavar='USER',
        help='username (typical e-mail address)')
    parser.add_option('-p', '--password', metavar='PASS',
        help='password')
    parser.add_option('-o', '--output', metavar='FILE',
        help='resulting vCard file name (standard output by default)')

    (options, args) = parser.parse_args()

    if len(args) != 1:
        parser.error('URL argument is required.')
    if not parse_ox_url(args[0]):
        parser.error('This is not a valid HTTP(S) URL.')

    user = options.user
    if not user:
        user = raw_input('Username: ')
    password = options.password
    if not password:
        password = getpass()

    output_filename = options.output


def main():
    init_options()

    conn = None
    out_fd = None

    if not output_filename or output_filename == '-':
        out_fd = sys.stdout
    else:
        out_fd = open(output_filename, 'wb')

    conn = None
    try:
        conn = OXConnection()
        folders = conn.list_contact_folders()

        for folder in folders:
            contacts_data = conn.get_contact_folder_contents(folder)
            for contact in contacts_data:
                vcard = make_vcard(contact)
                out_fd.write(vcard.encode('utf-8'))

    finally:
        if conn:
            conn.close()
        if out_fd and out_fd is not sys.stdout:
            out_fd.close()

if __name__ == '__main__':
    main()
