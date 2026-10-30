import os
import sys

import pytest
from docx import Document

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import FillApp  # noqa: E402


HEADERS = ['Nom', 'PRENOM', 'DATE DE NAISSANCE', 'NUMERO PASSPORT', 'NATIONALITE', 'Adresse', 'RESULTAT']
PEOPLE = [
    ['DOE', 'JOHN', '01/01/1990', 'P1234567', 'FRANCAISE', '10 rue de Paris', 'NEGATIF'],
    ['SMITH', 'JANE', '02/02/1985', 'P7654321', 'ALGERIENNE', '5 rue Didouche', 'POSITIF'],
]


def make_list(path, headers=HEADERS, rows=PEOPLE):
    doc = Document()
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    for col, header in enumerate(headers):
        table.cell(0, col).text = header
    for r, row in enumerate(rows, start=1):
        for col, value in enumerate(row):
            table.cell(r, col).text = f'  {value} '  # stray whitespace, as in real files
    doc.save(path)
    return path


def make_template(path):
    doc = Document()
    t1 = doc.add_table(rows=2, cols=3)
    t1.cell(0, 0).text = 'Full Name:'
    t1.cell(0, 1).text = 'Date Of Birth:'
    t1.cell(0, 2).text = 'Nationality:'
    t1.cell(1, 0).text = 'Passport:'
    t1.cell(1, 1).text = 'Address:'
    t2 = doc.add_table(rows=2, cols=5)
    t2.cell(1, 1).text = 'Test:'
    t2.cell(1, 3).text = 'Result:'
    t2.cell(1, 4).text = 'Date:'
    doc.save(path)
    return path


@pytest.fixture
def files(tmp_path):
    out = tmp_path / 'out'
    out.mkdir()
    return {
        'list': make_list(tmp_path / 'list.docx'),
        'template': make_template(tmp_path / 'template.docx'),
        'out': out,
    }


def test_read_exposant_list_maps_header_to_values(files):
    data = FillApp.read_exposant_list(files['list'])

    assert len(data) == 2
    assert data[0]['Nom'] == 'DOE'
    assert data[0]['PRENOM'] == 'JOHN'
    assert data[1]['RESULTAT'] == 'POSITIF'


def test_read_exposant_list_strips_whitespace(files):
    data = FillApp.read_exposant_list(files['list'])
    assert all(value == value.strip() for row in data for value in row.values())


def test_read_exposant_list_header_only(tmp_path):
    path = make_list(tmp_path / 'empty.docx', rows=[])
    assert FillApp.read_exposant_list(path) == []


def test_create_person_document_fills_template(files):
    person = dict(zip(HEADERS, PEOPLE[0]))
    FillApp.create_person_document(files['template'], files['out'], person)

    doc = Document(files['out'] / 'DOE JOHN.docx')
    t1, t2 = doc.tables[0], doc.tables[1]
    assert t1.cell(0, 0).text == 'Full Name: DOE JOHN'
    assert t1.cell(0, 1).text == 'Date Of Birth: 01/01/1990'
    assert t1.cell(0, 2).text == 'Nationality: FRANCAISE'
    assert t1.cell(1, 0).text == 'Passport: P1234567'
    assert t1.cell(1, 1).text == 'Address: 10 rue de Paris'
    assert t2.cell(1, 1).text == 'Test: pcr'
    assert t2.cell(1, 3).text == 'Result: NEGATIF'
    assert t2.cell(1, 4).text.startswith('Date: 20')


def test_create_person_document_optional_fields_missing(files):
    person = {'Nom': 'DOE', 'PRENOM': 'JOHN', 'DATE DE NAISSANCE': '01/01/1990', 'RESULTAT': 'NEGATIF'}
    FillApp.create_person_document(files['template'], files['out'], person)

    t1 = Document(files['out'] / 'DOE JOHN.docx').tables[0]
    assert t1.cell(1, 0).text == 'Passport:'
    assert t1.cell(1, 1).text == 'Address:'


def test_process_documents_creates_one_file_per_person(files):
    FillApp.process_documents(files['list'], files['template'], files['out'])

    assert sorted(os.listdir(files['out'])) == ['DOE JOHN.docx', 'SMITH JANE.docx']


def test_template_is_not_modified(files):
    before = files['template'].read_bytes()
    FillApp.process_documents(files['list'], files['template'], files['out'])
    assert files['template'].read_bytes() == before
