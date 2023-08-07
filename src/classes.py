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


def generate_mock_ticket():
    """
    Generate a mock ticket.

    Returns:
        CornTicket: The generated mock ticket.

    """
    return CornTicket(
        date=datetime.datetime(2021, 9, 17),
        ticket=852,
        gross=86600,
        tare=24640,
        mo=31.8,
        tw=50.4,
        driver="Dakota",
        truck="T-11",
    )
