import datetime


class CornTicket:
    def __init__(self, date, ticket, gross=0, tare=0, mo=0, tw=0, driver="", truck=""):
        self.date = date
        self.ticket = ticket
        self.gross = gross
        self.tare = tare
        self.mo = mo
        self.tw = tw
        self.net = gross - tare
        self.wetBu = self.net / 56
        self.dryBu = (
            (1 - ((self.mo - 15.5) * 0.012)) * self.net / 56
            if mo > 15.5
            else self.wetBu
        )
        self.driver = driver
        self.truck = truck


import datetime
import random


def generate_random_date(year):
    """
    Generate a random date within a given year.
    """
    # Generate a date within given year
    start_date = datetime.date(year, 1, 1)
    end_date = datetime.date(year, 12, 31)

    return start_date + (end_date - start_date) * random.random()


def generate_random_name():
    """
    Generate a random name from a list of common first names.
    """
    # A list of some common first names. You can modify this to suit your needs.
    names = [
        "James",
        "John",
        "Robert",
        "Mary",
        "Patricia",
        "Jennifer",
        "Linda",
        "Elizabeth",
        "William",
        "Richard",
        "David",
        "Susan",
        "Joseph",
        "Margaret",
        "Charles",
        "Thomas",
        "Christopher",
        "Daniel",
        "Matthew",
        "Sarah",
        "Jessica",
        "Emily",
        "Michael",
        "Jacob",
        "Mohamed",
        "Emma",
        "Joshua",
        "Amanda",
        "Andrew",
        "Brian",
        "Brandon",
    ]
    return random.choice(names)


def generate_random_feedlot_name():
    """
    Generates a random feedlot name by combining two typical feedlot naming elements.
    """
    name_elements = [
        "Ridgefield",
        "Prairie",
        "Cattle Co.",
        "Livestock",
        "Ranchers",
        "Farmers",
        "Grains",
        "Crops",
        "Holdings",
        "Valley",
        "Hilltop",
        "Green Pastures",
        "Harvest",
        "Angus",
        "Farming",
        "Acres",
    ]

    # Let's say a feedlot name consists of two elements.
    feedlot_name = " ".join(random.sample(name_elements, 2))

    return feedlot_name


def generate_random_field_number():
    """
    Generates a random field number.
    """
    return random.randint(1, 99)


def generate_mock_ticket():
    """
    Generate a mock CornTicket object.
    """
    return CornTicket(
        date=generate_random_date(2023),
        ticket=random.randint(1, 1000),
        gross=random.randint(50000, 99999),
        tare=random.randint(20000, 30000),
        mo=random.uniform(30.0, 40.0),  # Generate a random float
        tw=random.uniform(45.0, 55.0),  # Generate a random float
        driver=generate_random_name(),
        truck="T-" + str(random.randint(1, 20)),
    )


def generate_mock_tickets(num_tickets=2):
    """
    Generate a list of mock CornTicket objects based on the number of tickets provided.
    """
    return [generate_mock_ticket() for _ in range(num_tickets)]
