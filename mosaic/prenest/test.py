from mosaic import config
import xlwings

xlwings.Book(config['files']['workorder_template'])