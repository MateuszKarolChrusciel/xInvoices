import datetime
import os
import re

from common.data import regex_patterns as patterns


class Console:
    def __init__(self):
        raise ValueError

    @staticmethod
    def clear():
        os.system("cls")


class Invoice:
    def __init__(self, invoice_data: str):
        self._raw_content = invoice_data

        self._true_raw_content = self.get_true_raw_content()

        self._buyer_nip = self.get_buyer_nip()
        self._id = self.get_invoice_id()
        self._payment_date = self.get_payment_date()
        self._payment_method = self.get_payment_method()
        self._raw_prices = self._get_raw_prices()
        self._signing_date = self.get_signing_date()

        # print(f"Invoice #{self.id} initialized at {datetime.datetime.now()}")

    @property
    def correct(self):
        if self.id[0] == "K":
            return True
        else:
            return False

    @property
    def raw_content(self):
        return self._raw_content

    @property
    def true_raw_content(self):
        return self._true_raw_content

    @property
    def buyer_nip(self):
        return self._buyer_nip

    @property
    def gross_value(self):
        try:
            return self.raw_prices["gross"]
        except KeyError:
            return self.net_value + self.vat_value

    @property
    def id(self):
        return self._id

    @property
    def net_value(self):
        return self.raw_prices["net"]

    @property
    def payment_method(self):
        return self._payment_method

    @property
    def payment_date(self):
        return self._payment_date

    @property
    def raw_prices(self):
        return self._raw_prices

    @property
    def signing_date(self):
        return self._signing_date

    @property
    def vat_value(self):
        return self.raw_prices["vat"]

    def get_buyer_nip(self):
        results = []

        for pattern in patterns.nip_patterns:
            results.append(re.findall(pattern, self.raw_content))

        result = None

        for rl in results:
            if rl:
                for rs in rl:
                    if "773-156-23-77" in rs or "7731562377" in rs:
                        pass
                    else:
                        result = rs
            else:
                pass

        correct_result = ""

        try:
            for sign in result:
                if not sign.isdigit():
                    pass
                else:
                    correct_result = f"{correct_result}{sign}"
        except TypeError:
            raise NotImplementedError

        return f"{correct_result}"

    def get_true_raw_content(self):
        split_content = self.raw_content.split("\n")

        for counter, line in enumerate(split_content):
            if not line:
                split_content.pop(counter)
            else:
                while "\n" in line:
                    line -= "\n"

        return split_content

    def get_invoice_id(self):
        pattern = r"(?:numer)(?::*)(?:\s*)(\w*(/)*\w*((/)*\w*)*)"

        results = re.findall(pattern, self.raw_content)

        if len(results) == 1:
            result = results[0][0]
        else:
            raise NotImplementedError

        return result

    def get_payment_date(self):
        def find_line_with_pd(lines):
            line_w_pm = None

            for line in lines:
                if "termin płatności" in line.lower():
                    line_w_pm = line
                    break
                else:
                    pass

            return line_w_pm

        def get_raw_pd(line_w_pd):
            indexes = []

            for counter, sign in enumerate(line_w_pd):
                if sign.isdigit():
                    indexes.append(counter)

            pd_raw_str = line_w_pd[indexes[0]: None]

            try:
                y = int(pd_raw_str[0:4])
                m = int(pd_raw_str[5:7])
            except ValueError:
                raise NotImplementedError
            try:
                d = int(pd_raw_str[8:10])
            except ValueError or KeyError:
                d = int(pd_raw_str[8:None])

            return y, m, d

        line_with_payment_date = find_line_with_pd(self.true_raw_content)

        try:
            year, month, day = get_raw_pd(line_with_payment_date)
        except NotImplementedError:
            return None

        payment_date = datetime.date(year=year, month=month, day=day)

        return payment_date

    def get_payment_method(self):
        def find_line_with_pm(lines):
            line_w_pm = None

            for line in lines:
                if "płatność" in line.lower():
                    line_w_pm = line
                    break
                else:
                    pass

            if line_w_pm is None:
                return "nic"
            else:
                return line_w_pm

        line_with_pm = find_line_with_pm(self.true_raw_content)

        if "gotówka" in line_with_pm.lower():
            pm = "Gotówka"
        elif "przelew" in line_with_pm.lower():
            pm = "Przelew"
        elif "nic" in line_with_pm.lower():
            pm = ""
        else:
            raise ValueError

        return pm.upper()

    def get_signing_date(self):
        def find_line_w_sd(lines):
            line_w_sd = None
            found_line_w_sd = False

            while True:
                for line in lines:
                    if "wystawienia" in line.lower():
                        line_w_sd = line
                        found_line_w_sd = True
                        break
                if found_line_w_sd:
                    break
                else:
                    raise ValueError

            return line_w_sd

        def get_raw_sd(line_w_sd):
            indexes = []

            for counter, sign in enumerate(line_w_sd):
                if sign.isdigit():
                    indexes.append(counter)

            sd_raw_str = line_w_sd[indexes[0]: None]

            y = int(sd_raw_str[0:4])
            m = int(sd_raw_str[5:7])
            try:
                d = int(sd_raw_str[8:10])
            except ValueError or KeyError:
                d = int(sd_raw_str[8:None])

            return y, m, d

        line_with_signing_date = find_line_w_sd(self.true_raw_content)

        year, month, day = get_raw_sd(line_with_signing_date)

        date_when_signed = datetime.date(year=year, month=month, day=day)

        return date_when_signed

    def _get_raw_prices(self):
        regex = r"(\d{1,100}\s{1,100})*\d*[,]\d\d\s*PLN"

        matches = re.finditer(regex, self.raw_content, re.MULTILINE)

        finds = []

        for match in matches:
            finds.append(match.group())

        real_finds = []

        for find in finds:
            if find in real_finds:
                pass
            else:
                real_finds.append(find)

        for counter, find in enumerate(real_finds):
            if find.endswith(" PLN"):
                real_finds[counter] = find[:-4]

        prics = list()

        for find in real_finds:
            prics.append(float(find.replace(",", ".").replace(" ", "").replace("\n", "")))

        prices = dict()

        prices["gross"] = max(prics)
        prices["vat"] = min(prics)
        prices["net"] = prices["gross"] - prices["vat"]

        return prices


class Time:
    def __init__(self, hours: int, minutes: int, seconds: int, milliseconds: int):
        self.hours = hours
        self.minutes = minutes
        self.seconds = seconds
        self.milliseconds = milliseconds

    def __str__(self):
        if self.hours:
            return f"{self.hours}h, {self.minutes}min, {self.seconds}s and {self.milliseconds}ms ({self.secs}s)"
        elif self.minutes:
            return f"{self.minutes}min, {self.seconds}s and {self.milliseconds}ms ({self.secs}s)"
        elif self.seconds:
            return f"{self.seconds}s and {self.milliseconds}ms ({self.secs}s)"
        elif self.milliseconds:
            return f"{self.milliseconds}ms ({self.secs}s)"
        else:
            return "None"

    @property
    def secs(self):
        return self.hours * 3600 + self.minutes * 60 + self.seconds + self.milliseconds / 1000
