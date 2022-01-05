class ProcessingUtil:
    def __init__(self):
        self.aba_codes = ["97151", "97153", "97155", "97156"]
        self.eval_codes = ["96130", "96131", "96136", "96137"]

    def calc_check_content(self, filepath):
        file = open(filepath, "r")
        sum_aba, sum_eval = 0.0, 0.0
        charges = self.__breakdown_charges(file)
        for charge in charges:
            if charge[0] in self.aba_codes:
                sum_aba += float(charge[1])
            elif charge[0] in self.eval_codes:
                sum_eval += float(charge[1])

        return [sum_aba, sum_eval]

    @staticmethod
    def __breakdown_charges(file):
        lines, temp_result, result = [], [], []
        for line in file:
            if "Line Item" in line:
                temp = file.readline()
                temp2 = temp.strip()
                lines.append(temp2)

        for inner_list in lines:
            temp_list = inner_list.split(" ")
            temp_result.append(' '.join(temp_list).split())

        for item in temp_result:
            result.append([item[1], item[3]])

        return result
