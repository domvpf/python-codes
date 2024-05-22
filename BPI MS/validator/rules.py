from datetime import datetime
import re

class Rules:
    def check_csv(str_input, *args):
        return bool(str((str_input.lower())).endswith('.csv'))

    def check_xlsx(str_input, *args):
        return bool(str((str_input.lower())).endswith('.xlsx'))

    def check_header_fields(list_input1,list_input2):
        return list(filter(lambda item: not item in list_input2, list_input1))

    def header_hash_exists(header):
        return header[0] == 'H'

    def trailer_hash_exists(trailer):
        return trailer[0] == 'T'

    def at_least_one_record(num, *args):
        return num >= 1

    def check_header_hash(string, header_hash):
        return string.startswith(header_hash)

    def check_trailer_count(records, expected):
        return len(records) == expected

    def check_trailer_hash(records, expected):
        _sum = 0
        for record in records:
            whole, decimal = record['Net Premium'].split('.')
            _sum += (int(record['Policy No.'][-2:])* int(whole[-3:]))+ int(decimal[-2:])
        # print(_sum, expected)
        return _sum == expected

    def check_mandatory(row, fields):
        # print(row)
        # print(fields)
        return [k for k in fields if not row[k]]

    def check_type(header,data_type,*args):
        data_type = str(data_type)

        if header == '':
            return True
        elif data_type == 'A':
            # return bool(re.match('^[a-zA-Z0-9]+$',header))
            return True
        elif data_type == 'N':
            # return bool(re.match('^[0-9]+$', header))
            try:
                float(header)
                if header.find('E+')+1:
                    print(header)
                    raise ValueError
                return True
            except:
                return False
        elif data_type == 'D':
            # print(header,data_type)
            try:
                datetime.strptime(header, '%m/%d/%Y')
                return True
            except ValueError:
                return False

        elif data_type == '':
            return True
        elif data_type == 'nan':
            return True
        else:
            return False

    def check_length(field_length, length):
        # print(field_length,length)
        if length == '':
            return True
        # print(len(field_length), round(float(length)))
        return len(field_length) <= round(float(length)) and len(field_length) >= round(float(length))

    def check_range(d_value, range_min, range_max):
        if range_min and range_max:
            # print("min and max has value")
            # print("Min: ", range_min)
            # print("Max: ", range_max)
            if d_value < range_min:
                return False
            if d_value > range_max:
                return False
        elif range_min and range_max == '':
            # print("only min has value")
            # print("Min: ", range_min)
            if d_value < range_min:
                return False
        elif range_min == '' and range_max:
            # print("only max has value")
            # print("Max: ", range_max)
            if d_value > range_max:
                return False
        return True

    def check_contract_type(c_type, valid_contract_types):
        # print(c_type)
        # print(valid_contract_types)
        if c_type in valid_contract_types:
            return True
        else:
            return False
        # if c_type == 'odp':
        #     return True
        # elif c_type == 'otl':
        #     return True
        # elif c_type == 'opc':
        #    return True
        # else:
        #     # print("invalid")
        #     return False

    def check_condition_custom_rules(dict_dsf,cond_dsf):
        # print(dict_dsf,cond_dsf)
        bool_checker = True
        match_checker = True
        cond_counter = 0
        for i in range(0,len(dict_dsf)):
            dict_val = list(dict_dsf[i].values())
            # print(dict_val)
            if bool_checker == True:

                if dict_val[0] == dict_val[1]:
                    # print(dict_val[0],dict_val[1])
                    bool_checker = True
                    match_checker = True
                    # print("equal",match_checker)
                if cond_counter != len(dict_dsf)-1:
                    # print(cond_dsf[i])
                    dict_val_next = list(dict_dsf[i+1].values())
                    if cond_dsf[i] == 'or':
                        if dict_val_next[0] != dict_val_next[1]:
                            # print("not match")
                            bool_checker = True
                            match_checker = True
                    elif cond_dsf[i] == 'and':
                        # print(dict_val[0],dict_val[1])
                        if dict_val_next[0] != dict_val_next[1]:
                            # print("not match")
                            bool_checker = False
                            match_checker = False
                            # print(10 < 9)
                            return False
            cond_counter += 1


            if bool_checker == False:
                return False
            else:
                return True


class Template:
    def remove_emp_rows(ws_others):
        emp_row = [] # remove empty rows
        for index, row in enumerate(ws_others, start=1):
            empty = not any((cell.value for cell in row))
            if empty:
                emp_row.append(index)
        for row_index in reversed(emp_row):
            ws_others.delete_rows(row_index, 1)

        flag = False
        for row in ws_others.rows:
            for cell in row:
                if cell.value == 'POLICY NO':
                    ws_others.insert_rows(cell.row, amount=1)
                    flag=True
                    break
            if flag:
                break
            else:
                continue
