from check_registration import ECR
from check_registration import connecting_to_ecr
from check_registration import fr
from check_registration import logs_file_path
import datetime as dt


def main():

    ecr_on_test = ECR()

    if connecting_to_ecr():

        ecr_on_test.registration_report()

        ecr_on_test.open_session()

        ecr_on_test.cheque_without_position()

        ecr_on_test.cheque_with_different_tax_type()

        ecr_on_test.cheque_with_several_positions()

        ecr_on_test.close_session()



if __name__ == '__main__':
    print('Hello you in module main 2')
    main()