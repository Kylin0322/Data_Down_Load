
import selenium_chrome
import sys


def main(argv):

    web = selenium_chrome.webdriver_chrome()
    web.init()
    web.open_web()


if __name__ == '__main__':
    main(sys.argv)
