from docx import Document
import os


def create_renewal_letters(mail_list):
    try:
        if isinstance(mail_list, list):
            broker_name = input("\nEnter your full name: ")
            for row in mail_list:
                client_dict = generate_client_dict(row, broker_name)
                replace_str(client_dict)
        else:
            raise Exception("Argument has to be a list")

    except Exception as err:
        print(f"Unexpected {err=}, {type(err)=}")


def replace_str(client_dict):

    doc = Document("renewal_letter.docx")

    for p in doc.paragraphs:
        for key in client_dict:
            if client_dict[key] is None:
                client_dict[key] = ""
            inline = p.runs
            for j in range(len(inline)):
                if key in inline[j].text:
                    if key == 'c_year':
                        text = inline[j].text.replace(key, str(client_dict[key]))
                        inline[j].text = text
                    elif key == 'c_expiry':
                        text = inline[j].text.replace(key, client_dict[key].strftime("%d %b, %Y"))
                        inline[j].text = text
                    else:
                        text = inline[j].text.replace(key, client_dict[key])
                        inline[j].text = text

    if client_dict['owner_type'] == 'SINGLE':
        file_name = client_dict['license_plate']
    else:
        file_name = client_dict['license_plate'] + "_" + client_dict['owner_type']

    doc.save(rf'{os.path.abspath(os.curdir)}\TEMP\{file_name}.docx')


def generate_client_dict(client_info, broker_name):

    keys = ['last_name', 'first_name', 'address', 'city', 'postal_code',
            'c_expiry', 'c_year', 'c_make', 'c_model', 'license_plate', 'broker', 'f_name']
    client_dict = dict(zip(keys, client_info))

    # Check if it's owned by a company
    if client_info[10] is None and client_info[11] is None:
        comp_name = client_info[0].strip() + " " + client_info[1].strip()
        client_dict['last_name'] = comp_name
        client_dict['first_name'] = ''
        client_dict['f_name'] = comp_name
        client_dict['owner_type'] = 'COMPANY'

    # Check if there's more than one owner (different last name)
    elif ',' in client_info[0] or ',' in client_info[1]:
        client_info[0] = client_info[0].replace('&', '')
        client_info[1] = client_info[1].replace('&', '')
        first_owner = client_info[0].split(',')
        second_owner = client_info[1].split(',')
        client_dict['last_name'] = client_info[0].strip() + " & " + client_info[1].strip()
        client_dict['first_name'] = ''
        client_dict['f_name'] = first_owner[1].strip() + " & " + second_owner[1].strip()
        client_dict['owner_type'] = 'MULTIPLE'
    else:
        client_dict['last_name'] = client_dict['last_name'] + ","
        client_dict['f_name'] = client_dict['first_name']
        client_dict['owner_type'] = 'SINGLE'

    client_dict['broker'] = broker_name
    return client_dict

