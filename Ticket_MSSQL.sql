CREATE TABLE dbo.TicketType (
    TicketTypeID INT NOT NULL,
    Name         VARCHAR(50) NOT NULL,
    CONSTRAINT PK_TicketType PRIMARY KEY (TicketTypeID)
);

CREATE TABLE dbo.Ticket (
    TicketID      INT IDENTITY(1,1) NOT NULL, -- TicketID is Auto Generated
    Name          VARCHAR(100) NOT NULL,
    Description   VARCHAR(4000),
    IsDone        BIT NOT NULL,
    TicketTypeID  INT NOT NULL,
    DateCreated   DATE NOT NULL,
    CONSTRAINT PK_Ticket PRIMARY KEY (TicketID)
);

INSERT INTO dbo.TicketType (TicketTypeID, Name)
VALUES
    (1, 'Ticket Type 1'),
    (2, 'Ticket Type 2'),
    (3, 'Ticket Type 3');

INSERT INTO dbo.Ticket (
    Name,
    Description,
    IsDone,
    TicketTypeID,
    DateCreated
)
VALUES
    ('Test Ticket 1',  'Sample ticket 1',  0, 1, '2026-02-01'),
    ('Test Ticket 2',  'Sample ticket 2',  0, 2, '2026-02-01'),
    ('Test Ticket 3',  'Sample ticket 3',  1, 3, '2026-02-02'),
    ('Test Ticket 4',  'Sample ticket 4',  0, 1, '2026-02-02'),
    ('Test Ticket 5',  'Sample ticket 5',  1, 2, '2026-02-03'),
    ('Test Ticket 6',  'Sample ticket 6',  0, 3, '2026-02-03'),
    ('Test Ticket 7',  'Sample ticket 7',  0, 1, '2026-02-04'),
    ('Test Ticket 8',  'Sample ticket 8',  1, 2, '2026-02-04'),
    ('Test Ticket 9',  'Sample ticket 9',  0, 3, '2026-02-05'),
    ('Test Ticket 10', 'Sample ticket 10', 0, 1, '2026-02-05');
