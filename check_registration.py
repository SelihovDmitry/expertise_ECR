# Модуль регистрации чеков на ККТ
# пробивает чеки и возвращает чек с тегами из ФН

from win32com.client import Dispatch
import datetime as dt
import time

fr = Dispatch('Addin.DRvFR')
wait_cheque_timeout = 1


logs_file_path = 'result.txt'

def connecting_to_ecr():

    print('Начало работы, подключение к ККТ')

    with open(logs_file_path, 'w+') as log:  # w - открытие (если нет - создается) файла на запись
        log.write(f'{dt.datetime.now()}: Начало тестирования ККТ \n')
        fr.GetECRStatus()
        if fr.ResultCode == 0:
            print('Подключение к ККТ прошло успешно')
            fr.TableNumber = 18
            fr.RowNumber = 1
            fr.FieldNumber = 1
            fr.ReadTable()
            log.write(
                f'{dt.datetime.now()}: Подключение к ККТ з\н {fr.ValueOfFieldString}, код ошибки: {fr.resultcode}, {fr.resultcodedescription}\n')
            fr.GetDeviceMetrics()
            log.write(
                f'{dt.datetime.now()}: Модель ККТ {fr.UDescription}, прошивка {fr.ECRSoftVersion} от {dt.datetime.date(fr.ECRSoftDate)}\n')
            return True
        else:
            print(f'Подключение не удалось, код ошибки: {fr.resultcode}, {fr.resultcodedescription}')
            log.write(f'{dt.datetime.now()}: Подключение не удалось, код ошибки: {fr.resultcode}, {fr.resultcodedescription}\n')
            return False


class ECR:

    def __init__(self):
        pass

    def _get_cheque_from_fn(self):
        fr.FNGetStatus()
        fr.ShowTagNumber = True
        fr.FNGetDocumentAsString()
        return fr.StringForPrinting  # возвращаем документ из ФН в виде строки

    def registration_report(self):
        # метод получения отчета о регистрации
        print('Считываем отчет о регистрации из ФН')
        fr.DocumentNumber = 1
        fr.ShowTagNumber = True
        fr.FNGetDocumentAsString()
        time.sleep(wait_cheque_timeout)
        result = fr.StringForPrinting
        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последжнюю строку файла - для ДОзаписи вниз
            log.write(
                f'{dt.datetime.now()}: Отчет о регистрации, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
            log.write(f'Получен чек \n{result}\n')
        return result  # возвращаем документ из ФН в виде строки

    def open_session(self):
        # метод открытия смены
        print('Регистрируется чек открытия смены')
        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последжнюю строку файла - для ДОзаписи вниз
            fr.OpenSession()
            time.sleep(wait_cheque_timeout) # задержка - даем время на печать на всякий случай
            log.write(f'{dt.datetime.now()}: Открытие смены, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
            result = self._get_cheque_from_fn()
            log.write(f'Получен чек \n{result}')
            fr.Disconnect()
            return result

    def cheque_without_position(self):
        # Проверка формирования кассового чека без товарной позиции
        print('Регистрируется кассовый чек без товарной позиции')

        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последнюю строку файла - для записи вниз
            fr.GetECRStatus() # проверяем режим ККТ, если не 2 - выходим
            if fr.ECRMode == 2:
                fr.OpenCheck()
                fr.FNCloseCheckEx()
                time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                log.write(
                    f'{dt.datetime.now()}: Регистрация кассового чека без товарной позиции, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                if fr.resultcode == 0:
                    result = self._get_cheque_from_fn()
                    log.write(f'Получен чек \n{result}\n')
                    return result
                else:
                    print(f'Регистрация кассового чека без товарной позиции, код ошибки {fr.resultcode}, {fr.resultcodedescription}')
                    fr.CancelCheck()

                fr.Disconnect()
            else:
                return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

    def cheque_with_different_tax_type(self):
        # проверка формирования кассового чека с разными системами налогообложения
        tax_types = {1: 'ОСН',
                     2: 'УСН доход',
                     4: 'УСН доход - расход',
                     16: 'ЕСХН',
                     32: 'Патент'}
        for tax_type_value, tax_type_name in tax_types.items():
            print(f'Регистрируется кассовый чек c СНО {tax_type_name}')

            with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
                log.seek(0, 2)  # перемещаем курсор на последжнюю строку файла - для ДОзаписи вниз
                fr.GetECRStatus()  # проверяем режим ККТ, если не 2 - выходим
                if fr.ECRMode == 2:
                    fr.price = 1.11
                    fr.quantity = 1
                    fr.FNOperation()

                    fr.Summ1 = 100
                    fr.TaxType = tax_type_value
                    fr.FNCloseCheckEx()
                    time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                    log.write(
                        f'{dt.datetime.now()}: Регистрация чека c СНО {tax_type_name}, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                    if fr.resultcode == 0:
                        result = self._get_cheque_from_fn()
                        log.write(f'Получен чек \n{result}\n')
                        # return result
                    else:
                        print(
                            f'Регистрация чека c СНО {tax_type_name}, код ошибки {fr.resultcode}, {fr.resultcodedescription}')
                        fr.CancelCheck()

                    fr.Disconnect()

                else:
                    return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

    def cheque_with_several_positions(self):
        # проверка формирования кассового чека с несколькими товарными позициями
        print('Регистрируется кассовый чек с несколькими позициями')

        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последжнюю строку файла - для ДОзаписи вниз
            fr.GetECRStatus() # проверяем режим ККТ, если не 2 - выходим
            if fr.ECRMode == 2:
                fr.price = 1.11
                fr.quantity = 1
                fr.FNOperation()

                fr.price = 2.22
                fr.quantity = 2
                fr.FNOperation()

                fr.price = 3.33
                fr.quantity = 3
                fr.FNOperation()

                fr.Summ1 = 100
                fr.FNCloseCheckEx()
                time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                log.write(
                    f'{dt.datetime.now()}: Регистрация чека с несколькими позициями, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                if fr.resultcode == 0:
                    result = self._get_cheque_from_fn()
                    log.write(f'Получен чек \n{result}\n')
                    return result
                else:
                    print(
                        f'Регистрация чека с несколькими позициями, код ошибки {fr.resultcode}, {fr.resultcodedescription}')
                    fr.CancelCheck()
                fr.Disconnect()
            else:
                return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

    def cheque_with_different_tax(self):
        # проверка формирования кассового чека с разными налоговыми ставками
        taxes = {1: 'НДС 20%',
                 2: 'НДС 10%',
                 3: 'НДС 0%',
                 4: 'БЕЗ НДС',
                 5: 'НДС 20/120',
                 6: 'НДС 10/110'
                 }
        for tax_value, tax_name in taxes.items():
            print(f'Регистрируется кассовый чек c налоговой ставкой {tax_name}')

            with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
                log.seek(0, 2)  # перемещаем курсор на последнюю строку файла - для записи вниз
                fr.GetECRStatus()  # проверяем режим ККТ, если не 2 - выходим
                if fr.ECRMode == 2:
                    fr.price = 1.11
                    fr.quantity = 1
                    fr.tax1 = tax_value
                    fr.FNOperation()

                    fr.Summ1 = 100
                    fr.TaxType = 1
                    fr.FNCloseCheckEx()
                    time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                    log.write(
                        f'{dt.datetime.now()}: Регистрация чека c налоговой ставкой {tax_name}, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                    if fr.resultcode == 0:
                        result = self._get_cheque_from_fn()
                        log.write(f'Получен чек \n{result}\n')
                        # return result
                    else:
                        print(
                            f'Регистрация чека c налоговой ставкой {tax_name}, код ошибки {fr.resultcode}, {fr.resultcodedescription}')
                        fr.CancelCheck()

                    fr.Disconnect()

                else:
                    return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

    def cheque_with_all_tax(self):
        # проверка формирования кассового чека со всеми налоговыми ставками
        taxes = {1: 'НДС 20%',
                 2: 'НДС 10%',
                 3: 'НДС 0%',
                 4: 'БЕЗ НДС',
                 5: 'НДС 20/120',
                 6: 'НДС 10/110'
                 }

        print(f'Регистрируется кассовый чек со всеми налоговыми ставками')

        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последнюю строку файла - для записи вниз
            fr.GetECRStatus()  # проверяем режим ККТ, если не 2 - выходим
            if fr.ECRMode == 2:
                for tax_value, tax_name in taxes.items():
                    fr.StringForPrinting = f'товар с налоговой ставкой {tax_name}'
                    fr.price = 1.11
                    fr.quantity = 1
                    fr.tax1 = tax_value
                    fr.FNOperation()

                fr.Summ1 = 100
                fr.TaxType = 1
                fr.FNCloseCheckEx()
                time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                log.write(
                    f'{dt.datetime.now()}: Регистрация чека со всеми налоговыми ставками, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                if fr.resultcode == 0:
                    result = self._get_cheque_from_fn()
                    log.write(f'Получен чек \n{result}\n')
                    # return result
                else:
                    print(
                        f'Регистрация чека со всеми налоговыми ставками, код ошибки {fr.resultcode}, {fr.resultcodedescription}')
                    fr.CancelCheck()

                fr.Disconnect()

            else:
                return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

    def fn_operation_with_marking(self, price=1.11, quantity=1):
        # пробитие чека с маркировкой
        print('Регистрируется кассовый чек с маркировкой')
        self.price = price
        self.quantity = quantity

        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)
            fr.GetECRStatus()
            if fr.ECRMode == 2 or fr.ECRMode == 8:
                fr.price = self.price
                fr.quantity = self.quantity
                fr.PaymentTypeSign = 1  # ПризнакСпособаРасчета = Аванс
                fr.PaymentItemSign = 31
                fr.FNOperation()

                qr = "0102900021916404213Rfn-(uL4hLHv\x1D91EE06\x1D92ZL1qUSqxS/jylFxi1Sp/HouC05T7FqUi34uslMAoDc8="
                fr.BarCode = qr
                fr.ItemStatus = 1
                fr.FNSendItemBarcode()
                time.sleep(wait_cheque_timeout)
                print(f'Передача марки, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                log.write(
                    f'{dt.datetime.now()}: Передача марки, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')

                fr.TagNumber = 1262 # ИД. ФОИВ
                fr.TagType = 7
                fr.TagValueStr = "001"
                fr.FNSendTagOperation()
                fr.TagNumber = 1263 # ДАТА ДОК. ОСН.
                fr.TagType = 7
                fr.TagValueStr = "13.05.2024"
                fr.FNSendTagOperation()
                fr.TagNumber = 1264 # НОМЕР ДОК. ОСН.
                fr.TagType = 7
                fr.TagValueStr = "22"
                fr.FNSendTagOperation()
                fr.TagNumber = 1265 # ЗНАЧ. ОТР. РЕКВ.
                fr.TagType = 7
                fr.TagValueStr = "ЗНАЧ. ОТР. РЕКВ."
                fr.FNSendTagOperation()

                fr.Summ1 = 100
                fr.FNCloseCheckEx()
                time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                log.write(
                    f'{dt.datetime.now()}: Регистрация чека с маркировкой, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                result = self._get_cheque_from_fn()
                log.write(f'Получен чек \n{result}')
                fr.Disconnect()
                return result
            else:
                return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

    def close_session(self):
        # метод закрытия смены
        print('Регистрируется чек закрытия смены')
        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последжнюю строку файла - для ДОзаписи вниз
            fr.FNCloseSession()
            time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
            log.write(f'{dt.datetime.now()}: Закрытие смены, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
            result = self._get_cheque_from_fn()
            log.write(f'Получен чек \n{result}')
            fr.Disconnect()
            return result

    def calculation_state_report(self):
        # метод снятия отчета о состоянии расчетов
        print('Регистрируется документ отчета о состоянии расчетов')
        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последжнюю строку файла - для ДОзаписи вниз
            fr.FNBuildCalculationStateReport()
            time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
            log.write(f'{dt.datetime.now()}: Отчет о состоянии расчетов, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
            result = self._get_cheque_from_fn()
            log.write(f'Получен чек \n{result}\n')
            fr.Disconnect()
            return result

    def cheque_correction(self, price=1.11, quantity=1):
        # метод пробития чека на ККТ, возвращает текст чека из ФН
        print('Регистрируется чек коррекции')
        self.price = price
        self.quantity = quantity

        with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
            log.seek(0, 2)  # перемещаем курсор на последжнюю строку файла - для ДОзаписи вниз
            fr.GetECRStatus()  # проверяем режим ККТ, если не 2 - выходим
            if fr.ECRMode == 2:
                fr.CheckType = 0
                fr.FNOpenCheckCorrection()
                fr.price = self.price
                fr.quantity = self.quantity
                fr.FNOperation()

                fr.Summ1 = 100
                fr.FNCloseCheckEx()
                time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                log.write(
                    f'{dt.datetime.now()}: Регистрация чека коррекции, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                result = self._get_cheque_from_fn()
                log.write(f'Получен чек \n{result}')
                fr.Disconnect()
                return result
            else:
                return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')



if __name__ == '__main__':
    print('Hello you in module check_registration')
    ShtrihZnak = ECR()
    # print
    ShtrihZnak.cheque_with_all_tax()