import tableauserverclient as TSC
import pandas as pd
import numpy as np
import os
import re
from PyQt5 import QtCore, QtWidgets


class Tableau_PDF(object):
    def __init__(self):
        self.server = TSC.Server('https://prod-apnortheast-a.online.tableau.com')
        self.server.version = '3.9'
        self.dashboard_filter_file_name = 'dashboard_filters.csv'
        self.site_name = 'rexwal2'
        self.create_allowed_view_mapping()

    def create_allowed_view_mapping(self):
        df = pd.read_csv(self.dashboard_filter_file_name)
        df = df.replace(np.nan, '', regex=True)
        df.columns = df.columns.str.strip()
        dashboard_names = df['View'].to_list()
        dashboard_names = [x.strip(' ') for x in dashboard_names]
        filter_csv_file_list = df['Filter File'].to_list()
        filter_csv_file_list = [x.strip(' ') for x in filter_csv_file_list]
        self.views_filter_mapping = dict(zip(dashboard_names, filter_csv_file_list))

    def setup_ui(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(594, 610)
        self.email_label = QtWidgets.QLabel(Dialog)
        self.email_label.setGeometry(QtCore.QRect(30, 80, 121, 31))
        self.email_label.setObjectName("email_label")
        self.password_label = QtWidgets.QLabel(Dialog)
        self.password_label.setGeometry(QtCore.QRect(30, 160, 121, 31))
        self.password_label.setObjectName("password_label")
        self.generate_pdf = QtWidgets.QPushButton('generate_pdf', Dialog)
        self.generate_pdf.setGeometry(QtCore.QRect(170, 230, 131, 51))
        self.generate_pdf.clicked.connect(self.button_click)
        self.email = QtWidgets.QLineEdit(Dialog)
        self.email.setGeometry(QtCore.QRect(150, 80, 351, 31))
        self.email.setObjectName("email")
        self.password = QtWidgets.QLineEdit(Dialog)
        self.password.setGeometry(QtCore.QRect(150, 160, 351, 31))
        self.password.setObjectName("password")
        self.header = QtWidgets.QLabel(Dialog)
        self.header.setGeometry(QtCore.QRect(70, 20, 371, 41))
        self.header.setObjectName("header")

        self.create_list_items_logs(Dialog)
        self.retranslate_ui(Dialog)

        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def button_click(self):
        self.email_text = self.email.text()
        self.password_text = self.password.text()
        self.delete_all_logs()

        if not self.email_text or not self.password_text:
            self.add_item_logs("Email or password is not filled")
            return
        if not self.check_email(self.email_text):
            return

        self.tableau_login()

    def create_list_items_logs(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        self.count = 1
        self.logs_list = QtWidgets.QListWidget(Dialog)
        self.logs_list.setGeometry(QtCore.QRect(25, 301, 531, 241))
        self.logs_list.setStyleSheet("background-color: rgb(0, 0, 0);\n"
                                      "color: rgb(253, 250, 255);")
        self.logs_list.setObjectName("logs_list")
        item = QtWidgets.QListWidgetItem()
        self.logs_list.addItem(item)
        __sortingEnabled = self.logs_list.isSortingEnabled()
        self.logs_list.setAutoScroll(True)
        self.logs_list.setSortingEnabled(False)
        self.logs_list.setAutoScrollMargin(2000)
        item = self.logs_list.item(0)
        item.setText(_translate("Dialog", "Logs will be printed here"))
        self.logs_list.setSortingEnabled(__sortingEnabled)

    def add_item_logs(self, log_text):
        _translate = QtCore.QCoreApplication.translate
        item = QtWidgets.QListWidgetItem()
        item.setText(_translate("Dialog", log_text))
        self.logs_list.addItem(item)
        self.count += 1

    def delete_all_logs(self):
        for i in range(self.count):
            self.logs_list.takeItem(0)
        self.add_item_logs("Logs will be printed here")

    def retranslate_ui(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.email_label.setText(_translate("Dialog", "Tableau Email Id"))
        self.password_label.setText(_translate("Dialog", "Tableau Password"))
        self.generate_pdf.setText(_translate("Dialog", "Genrate PDFs"))
        self.header.setText(_translate("Dialog",
                                       "<html><head/><body><p><span style=\" font-size:18pt; font-weight:600;\">Tableau Dashboards PDF Downloader</span></p></body></html>"))

    def tableau_login(self):
        self.tableau_auth = TSC.TableauAuth(self.email_text, self.password_text, self.site_name)
        try:
            with self.server.auth.sign_in(self.tableau_auth):
                view_mapping = {}
                all_views, pagination_item = self.server.views.get()

                for view in all_views:
                    view_mapping.update({view.name: view})

                self.add_item_logs('Available view in tableau')
                self.add_item_logs(' '.join([view.name for view in all_views]))
                allowed_views = self.views_filter_mapping.keys()

                for view_name in allowed_views:
                    view_item = view_mapping.get(view_name)
                    self.save_all_pdf(view_item, view_name, self.views_filter_mapping.get(view_name))
            self.add_item_logs("Done now check the files")
        except Exception as e:
            self.add_item_logs("Not able to login")

    def save_all_pdf(self, view_item, view_name, filter_file_name):
        if view_item is None:
            return

        filter_values, filter_names = self.get_filter_values(filter_file_name)

        for value in filter_values:
            pdf_req_option = TSC.PDFRequestOptions(page_type=TSC.PDFRequestOptions.PageType.A3,
                                                   orientation=TSC.PDFRequestOptions.Orientation.Portrait,
                                                   maxage=1)
            file_name = ""
            for i in range(len(filter_names)):
                filter_name = filter_names[i]
                file_name = file_name + '-' + str(value[filter_name.get('pos')].strip())
                pdf_req_option.vf(filter_name.get('name'), value[filter_name.get('pos')].strip())

            self.server.views.populate_pdf(view_item, pdf_req_option)

            if not os.path.exists('pdf/' + view_name):
                os.makedirs('pdf/' + view_name)

            with open('./pdf/' + view_name + '/file-' + file_name + '.pdf', 'wb') as f:
                self.add_item_logs("Saving file {} for view {} ".format(view_name + file_name, view_name))
                f.write(view_item.pdf)

    def get_filter_values(self, filter_file_name):
        df = pd.read_csv(filter_file_name)
        df = df.replace(np.nan, '', regex=True)
        filter_names = []
        items = df.columns.values.tolist()

        for pos in range(len(items)):
            filter_name = items[pos]

            if "Unnamed" in filter_name:
                continue
            filter_names.append({'name': filter_name.strip(), 'pos': pos})

        return df.values.tolist(), filter_names

    def check_email(self, email):
        regex = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}$'
        if (re.search(regex, email)):
            return True
        self.add_item_logs("Invalid Email")
        return False


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Tableau_PDF()
    ui.setup_ui(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
