from lxml import etree
from random import shuffle
import requests
from bs4 import BeautifulSoup
import openpyxl
import argparse


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--output_path',
        help='path to save xlsx file',
        default='coursera.xlsx'
    )
    parser.add_argument(
        '--amount_of_courses',
        type=int,
        help='amount of courses to get info',
        default=20
    )
    parser.add_argument(
        '--lxml_url',
        help='url with courses',
        default='https://www.coursera.org/sitemap~www~courses.xml'
    )
    return parser.parse_args()


def fetch_coursera_lxml_feed(lxml_url):
    return requests.get(lxml_url).text


def get_url_courses_list(lxml_page, amount):
    courses_url_list = []
    xml = bytes(bytearray(lxml_page, encoding='utf-8'))
    tree = etree.fromstring(xml)
    for child in tree.getiterator():
        if 'https' in child.text:
            courses_url_list.append(child.text)
    shuffle(courses_url_list)
    return courses_url_list[:amount]


def fetch_course_page(course_url):
    return requests.get(course_url).text


def get_course_info(course_html_code):
    soup = BeautifulSoup(course_html_code, 'html5lib')
    name_of_course = soup.find('h1').text
    language = soup.find('div', 'rc-Language').text
    try:
        user_rating = soup.find('div', 'ratings-text').text
    except AttributeError:
        user_rating = None
    try:
        weeks = len(soup.find('div', 'rc-WeekView'))
    except TypeError:
        weeks = None
    start_date = soup.find('div', 'startdate').text
    return {
            'name_of_course': name_of_course,
            'weeks': weeks,
            'language': language,
            'user_rating': user_rating,
            'start_date': start_date,
    }


def output_courses_info_to_workbook(course_list):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append([
        'Url',
        'Name of course',
        'Weeks',
        'Language',
        'User Rating',
        'Start Date'
    ])
    for course_url in course_list:
        course_info = get_course_info(fetch_course_page(course_url))
        worksheet.append([
            course_url,
            course_info['name_of_course'],
            course_info['weeks'],
            course_info['language'],
            course_info['user_rating'],
            course_info['start_date']
        ])
    return workbook


if __name__ == '__main__':
    arguments = parse_args()
    try:
        output_courses_info_to_workbook(
            get_url_courses_list(
                fetch_coursera_lxml_feed(arguments.lxml_url),
                arguments.amount_of_courses)
        ).save(arguments.output_path)
    except FileNotFoundError:
        print('Download file from link in README')
