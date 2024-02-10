from textwrap import dedent


def messege_generator(data: dict, mentor: str = "Lakshmikantha G C", sem: int = 5):
    """ Returns the messege to be sent to the parent
        data: dict - The student data (default: None)
        mentor: str - The mentor name (default: "Lakshmikantha G C")
        sem: int - The semester number (1, 2, 3, etc.) """

    header = dedent(f"""
    Dear Parent,%0A
      Kindly find CIE {data['internal']} Marks:%0A
      USN No: {data['usn']}%0A
      Student Name : {data['name']}%0A
    """)[1:]

    body = ""
    for key in data['marks'].keys():
        body += f"  {key} : {data['marks'][key]}%0A\n"

    footer = dedent(
        f"""Regards, %0A\n  {mentor} Class Teacher %0A\n  {sem} Sem AIML""")

    return header + body + footer


# For testing
if __name__ == "__main__":
    data = {'internal': 1, 'marks': {'21MAT41A': '35 / 40', '21AML42': '17/30', '21AML43': '23/30', '21AML44': '21/40', '21AML45': '30/40',
                                     '21KBK46': 'NR/40', '21KSK46': '39/40'}, 'usn': '1GA22AI404', 'name': 'SHIVA GURURAJ  BAGEWADI', 'phone_number': '8050158461'}
    print(messege_generator(data))
