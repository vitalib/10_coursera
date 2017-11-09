import argparse
import requests
import re
import random
import bs4
import openpyxl


def get_xml_data():
    coursera_url = "https://www.coursera.org/sitemap~www~courses.xml"
    response = requests.get(coursera_url)
    return response.text


def get_courses_list(courses_quantity):
    soup = bs4.BeautifulSoup(get_xml_data(), 'xml')
    courses_list = [course.string for course in soup.find_all('loc')]
    return random.sample(courses_list, courses_quantity)


def get_course_info(course_html, verbose):
    soup = bs4.BeautifulSoup(course_html, 'html.parser')
    course_parse_data = {
        'Name': ('h1',),
        'Language': ('div', {'class': 'rc-Language'}),
        'Starting date': (
                     'div',
                     {'class': 'startdate rc-StartDateString caption-text'},
        ),
        'Rating': ('div', {'class': 'ratings-text bt3-hidden-xs'}),
    }
    course_text_info = dict.fromkeys(course_parse_data.keys(), None)
    course_text_info['Duration'] = None
    for field in course_parse_data:
        parse_data = soup.find(*course_parse_data[field])
        if parse_data:
            course_text_info[field] = parse_data.get_text()
    if course_text_info['Rating']:
        course_text_info['Rating'] = course_text_info['Rating'][20:23]
    duration = soup.find(string=re.compile('weeks of study'))
    if duration:
        course_text_info['Duration'] = duration.string
    return course_text_info


def output_courses_info_to_xlsx(courses_info_list):
    columns_names = ('Name', 'Language',
                     'Starting date', 'Duration',
                     'Rating',
                     )
    work_book = openpyxl.Workbook()
    work_sheet = work_book.active
    work_sheet.append(columns_names)
    for course_info in courses_info_list:
        course_data = [
            course_info[column] for column in columns_names]
        work_sheet.append(course_data)
    return work_book


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('filepath',
                        help='path to output xlsx file'
                        )
    parser.add_argument('-v', '--verbose', action='store_true')
    return parser.parse_args()


if __name__ == '__main__':
    args = get_args()
    courses_url_list = get_courses_list(courses_quantity=20)
    courses_info_list = []
    for course_url in courses_url_list:
        if args.verbose:
            print('Handling {} ...'.format(course_url))
        response = requests.get(course_url)
        courses_info_list.append(get_course_info(response.text,
                                                 args.verbose,
                                                 )
                                 )
        if args.verbose:
            print('Ok')
    work_book_xlsx = output_courses_info_to_xlsx(courses_info_list)
    work_book_xlsx.save(args.filepath)
    if args.verbose:
        print('Job is done. Result is stored in {}'.format(args.filepath))
