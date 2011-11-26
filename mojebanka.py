#!/usr/bin/python
# -*- coding: utf-8 -*-

import re
import codecs
import time

class ParserInvalidTransaction(Exception): pass

# Example of transaction values:
#
# - account   ... 94-60576642/8060
# - var_sym   ... 8205868823
# - price     ... -1500,00 CZK
# - date1     ... 08.09.2011
# - type      ... Úhrada
# - const_sym ... 0
# - date2     ... 08.09.2011
# - trans_id  ... 000-08092011 005-005-001596020
# - spec_s    ... 0
# - date3     ... 08.09.2011
# - desc1     ... STAVEBNI SPORENI - BURINKA
# - desc2     ... NA   AC-0000940060576642
# - desc3     ... Úhrada do jiné banky
# - desc4     ...

def mojebanka_txt_parse(content):
    transactions = []
    cells = content.split('________________________________________________________________________________')
    re_skip = re.compile(u".*(ČÍSLO ÚČTU : |Obrat na vrub|Číslo protiúčtu                VS|Transakční historie|Za období      od).*", re.DOTALL)
    re_parse = re.compile(u"(?P<account>\\d*/\\d{4})[ ]*(?P<var_sym>\\d*)[ ]*(?P<price>-?\\+?\\d+,\\d{2} CZK)[ ]*(?P<date1>\\d{2}\\.\\d{2}.\\d{4})\\r?\\n"
                          u"(?P<type>|Úhrada|Inkaso)[ ]*(?P<const_sym>\\d+)[ ]*(?P<date2>\\d{2}\\.\\d{2}.\\d{4})\\r?\\n"
                          u"(?P<trans_id>\\d[0-9A-Z -]{14,31})[ ]*(?P<date3>\\d+)[ ]*(\\d{2}\\.\\d{2}.\\d{4})\\r?\\n"
                          u"Popis příkazce[ ]*(?P<desc1>.+)\\r?\\nPopis pro příjemce[ ]*(?P<desc2>.+)\\r?\\n"
                          u"Systémový popis[ ]*(?P<desc3>.+)")
    try:
        for cell in cells:
            cell = cell.strip(" \r\n\t")
            if not cell or re_skip.match(cell):
                continue
            else:
                r = re_parse.search(cell)
                if not r:
                    raise ParserInvalidTransaction();
                else:
                    # Get parsed keys
                    tr = r.groupdict()

                    # Desc 4 needs to be parsed separately
                    start = cell.find(u"Zpráva pro příjemce")
                    tr['desc4'] = re.sub('\n|\r|\t', ' ', cell[start + 19:]) if start != -1 else ""

                    # Cleanup the parsed strings
                    for key in tr.iterkeys():
                        tr[key] = tr[key].strip()
                    # Convert dates
                    for key in ('date1', 'date2', 'date3'):
                        tr[key] = time.strptime(tr[key], '%d.%m.%Y') if len(tr[key]) == 10 else None
                    # Cleanup descriptions 
                    for key in ('desc1', 'desc2', 'desc3', 'desc4'):
                        tr[key] = re.sub('[ ]{2,}', ' ', tr[key]) if tr[key] else ""

                    transactions.append(tr)

    except (ParserInvalidTransaction) as err:
        print "\n\n=========================\nUnknown transaction!\n" + cell;

    return transactions


def mojebanka_to_cvs(transactions):
    """ Save array of transactions to CVS file. """
    try:
        filename = date_filename('cvs')
        fout = open(filename, 'w+')
        columns = ['date3', 'type', 'account', 'price', 'var_sym', 'desc1', 'desc2', 'desc3', 'desc4']
        fout.write("\t" . join(columns))

        for tr in transactions:
            data = []
            for col in columns:
                data.append(tr[col]);
            fout.write("\t".join(data))
    finally:
        fout.close()


def mojebanka_to_qif(transactions):
    """ Save array of transactions to QIF file. """
    filename = date_filename('qif')
    qif_file = codecs.open(filename, 'w+', 'utf-8')
    qif_file.write("!Type:Bank\n")

    for tr in transactions:
        amount = re.sub('\+?(-?)(\d+),(\d{2}) CZK', '\g<1>\g<2>.\g<3>', tr['price'])
        if tr['account'] == '\0100':
            payee = 'KB'
        else:
            payee = tr['account']

        if tr['var_sym']:
          payee += ' ' + tr['var_sym']

        data = '';
        data += 'D' + time.strftime('%d/%m/%Y', tr['date1']) + "\n"
        data += 'T' + number_format (amount, 2) + "\n"
        data += 'P' + payee + "\n"
        data += 'M' + tr['desc1'] + ' ' + tr['desc2'] + ' ' + tr['desc3'] + ' ' + tr['desc4'] + "\n"
        data += "^\n"
        qif_file.write(data)

    qif_file.close()



def date_filename(ext):
    return "mojebanka_export_" + time.strftime("%Y-%d-%m-%H-%M-%S") + "." + ext


def number_format(number, decimals = 0, dec_point = '.', thousands_sep = ','):
    """
    This filter allows you to format numbers like PHP's number_format function. 
    http://djangosnippets.org/snippets/682/ 
    """
    try:
        number = round(float(number), decimals)
    except ValueError:
        return number
    neg = number < 0
    integer, fractional = str(abs(number)).split('.')
    m = len(integer) % 3
    if m:
        parts = [integer[:m]]
    else:
        parts = []

    parts.extend([integer[m + t:m + t + 3] for t in xrange(0, len(integer[m:]), 3)])

    if decimals:
        return '%s%s%s%s' % (
            neg and '-' or '',
            thousands_sep.join(parts),
            dec_point,
            fractional.ljust(decimals, '0')[:decimals]
        )
    else:
        return '%s%s' % (neg and '-' or '', thousands_sep.join(parts))


################################################################################
## MAIN LOOP 
if __name__ == '__main__':
    import glob
    from optparse import OptionParser

    usage = "%prog [moznosti] soubor\n\nPriklad: %prog -f=csv *.txt"
    optparser = OptionParser(usage = usage)
    optparser.add_option("-f", "--format", dest = "format", default = "qif", help = "Cilovy format: qif, cvs")

    (options, args) = optparser.parse_args()

    files = []
    for file in args:
        new_files = glob.glob(file);
        for new_file in new_files:
            if new_file not in files:
                files.append(new_file)

    if len(files):
        files = list(set(files))
    else:
        optparser.print_help();

    for file in files:            
        fin = codecs.open(file, 'r', 'cp1250')
        content = fin.read()
        transactions = mojebanka_txt_parse(content)
        if options.format == 'cvs':        
            mojebanka_to_cvs(transactions)
        else:
            mojebanka_to_qif(transactions)
        fin.close()
