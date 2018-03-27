import xml.etree.ElementTree as et
from random import shuffle
import requests
from bs4 import BeautifulSoup
import openpyxl
import argparse


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--output_path',
        type=str,
        help='path to save xlsx file',
        default='coursera.xlsx'
    )
    parser.add_argument(
        '--amount_of_courses',
        type=int,
        help='amount of courses to get info',
        default=20
    )
    return parser.parse_args()


def get_courses_list(amount):
    courses_list = []
    tree = et.parse('sitemap_www_courses.xml')
    root = tree.getroot()
    for child in root.iter():
        if 'https' in child.text:
            courses_list.append(child.text)
    shuffle(courses_list)
    return courses_list[:amount]


def get_course_info(course_url):
    html_page = requests.get(course_url).text
    soup = BeautifulSoup(html_page, "html5lib")
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
            'url': course_url,
            'name': name_of_course,
            'weeks': weeks,
            'language': language,
            'user_rating': user_rating,
            'start_date': start_date
            }


def output_courses_info_to_xlsx(course_list, file_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append([
        'Url',
        'Name',
        'Weeks',
        'Language',
        'User Rating',
        'Start Date'
    ])
    for course_url in course_list:
        course_info = get_course_info(course_url)
        worksheet.append([
            course_info['url'],
            course_info['name'],
            course_info['weeks'],
            course_info['language'],
            course_info['user_rating'],
            course_info['start_date']
        ])
    workbook.save(file_path)


if __name__ == '__main__':
    arguments = parse_args()
    output_courses_info_to_xlsx(
        get_courses_list(arguments.amount_of_courses),
        arguments.output_path
    )
