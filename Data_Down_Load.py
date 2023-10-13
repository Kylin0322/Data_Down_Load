
import selenium_chrome
import sys


def main(argv):

    web = selenium_chrome.webdriver_chrome()
    web.init()
    print('删除旧文件', web.org_file_name_full)
    web.delete_file(web.org_file_name_full)

    web.open_web()
    print('拷贝文件到指定目录')
    print(web.org_file_name_full)
    print(web.des_path)
    web.move_file(file_name=web.org_file_name_full, folder=web.des_path)


if __name__ == '__main__':
    main(sys.argv)
