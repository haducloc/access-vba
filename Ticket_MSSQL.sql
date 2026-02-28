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

CREATE TABLE [dbo].[TicketComment](
	[TicketCommentID] [int] IDENTITY(1,1) NOT NULL,
	[Comment] [varchar](255) NOT NULL,
	[TicketID] [int] NOT NULL,
	CONSTRAINT PK_TicketComment PRIMARY KEY (TicketCommentID)
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

INSERT INTO dbo.TicketComment (Comment, TicketID) VALUES
-- Ticket 1
('Initial issue reported by user.', 1),
('Support team investigating the problem.', 1),
('Issue resolved and confirmed by user.', 1),

-- Ticket 2
('Login failure reported.', 2),
('Password reset instructions sent.', 2),
('User successfully logged in after reset.', 2),

-- Ticket 3
('Application crashes on startup.', 3),
('Logs requested from user.', 3),
('Patch deployed to fix startup crash.', 3),

-- Ticket 4
('Unable to generate report.', 4),
('Database connection timeout identified.', 4),
('Connection settings updated and verified.', 4),

-- Ticket 5
('Email notifications not working.', 5),
('SMTP configuration reviewed.', 5),
('Email notifications restored.', 5),

-- Ticket 6
('Performance degradation noticed.', 6),
('High CPU usage detected on server.', 6),
('Server restarted and performance normalized.', 6),

-- Ticket 7
('Error message displayed during checkout.', 7),
('Payment gateway timeout confirmed.', 7),
('Retry logic implemented and tested.', 7),

-- Ticket 8
('Data export not downloading.', 8),
('File permission issue discovered.', 8),
('Permissions corrected and export successful.', 8),

-- Ticket 9
('UI layout broken on mobile.', 9),
('CSS conflict identified.', 9),
('Styles updated and verified on mobile devices.', 9),

-- Ticket 10
('User account locked unexpectedly.', 10),
('Security logs reviewed.', 10),
('Account unlocked and user notified.', 10);