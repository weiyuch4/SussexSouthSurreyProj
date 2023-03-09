import SussexData
import renewal

if __name__ == "__main__":
    x = SussexData.SussexData()
    lst = x.generate_mail_list()
    renewal.create_renewal_letters(lst)

