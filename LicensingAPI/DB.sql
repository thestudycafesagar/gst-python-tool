/****** Object:  Table [dbo].[UserDevices] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UserDevices]') AND type = N'U')
BEGIN
CREATE TABLE [dbo].[UserDevices](
	[DeviceID] [int] IDENTITY(1,1) NOT NULL,
	[UserID] [int] NULL,
	[HardwareID] [nvarchar](100) NOT NULL,
	[DeviceName] [nvarchar](100) NULL,
	[AddedAt] [datetime] NULL,
PRIMARY KEY CLUSTERED
(
	[DeviceID] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UQ_User_Hardware] UNIQUE NONCLUSTERED
(
	[UserID] ASC,
	[HardwareID] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[Users] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Users]') AND type = N'U')
BEGIN
CREATE TABLE [dbo].[Users](
	[UserID] [int] IDENTITY(1,1) NOT NULL,
	[Username] [nvarchar](50) NOT NULL,
	[Password] [nvarchar](100) NOT NULL,
	[IsPremium] [bit] NULL,
	[IsActive] [bit] NULL,
	[AllowedDeviceLimit] [int] NULL,
	[TrialExpiryDate] [datetime] NULL,
	[LastLogin] [datetime] NULL,
	[CreatedAt] [datetime] NULL,
PRIMARY KEY CLUSTERED
(
	[UserID] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
UNIQUE NONCLUSTERED
(
	[Username] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM sys.default_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[UserDevices]') AND COL_NAME(parent_object_id, parent_column_id) = 'AddedAt')
    ALTER TABLE [dbo].[UserDevices] ADD DEFAULT (getdate()) FOR [AddedAt]
GO
IF NOT EXISTS (SELECT * FROM sys.default_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[Users]') AND COL_NAME(parent_object_id, parent_column_id) = 'IsPremium')
    ALTER TABLE [dbo].[Users] ADD DEFAULT ((0)) FOR [IsPremium]
GO
IF NOT EXISTS (SELECT * FROM sys.default_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[Users]') AND COL_NAME(parent_object_id, parent_column_id) = 'IsActive')
    ALTER TABLE [dbo].[Users] ADD DEFAULT ((1)) FOR [IsActive]
GO
IF NOT EXISTS (SELECT * FROM sys.default_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[Users]') AND COL_NAME(parent_object_id, parent_column_id) = 'AllowedDeviceLimit')
    ALTER TABLE [dbo].[Users] ADD DEFAULT ((1)) FOR [AllowedDeviceLimit]
GO
IF NOT EXISTS (SELECT * FROM sys.default_constraints WHERE parent_object_id = OBJECT_ID(N'[dbo].[Users]') AND COL_NAME(parent_object_id, parent_column_id) = 'CreatedAt')
    ALTER TABLE [dbo].[Users] ADD DEFAULT (getdate()) FOR [CreatedAt]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE parent_object_id = OBJECT_ID(N'[dbo].[UserDevices]') AND referenced_object_id = OBJECT_ID(N'[dbo].[Users]'))
    ALTER TABLE [dbo].[UserDevices] WITH CHECK ADD FOREIGN KEY([UserID]) REFERENCES [dbo].[Users] ([UserID])
GO
/****** Object:  StoredProcedure [dbo].[sp_AuthenticateWithDeviceManagement] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Table Name: Users, UserDevices
CREATE OR ALTER PROCEDURE [dbo].[sp_AuthenticateWithDeviceManagement]
    @InEmail NVARCHAR(100),
    @InPassword NVARCHAR(100),
    @InHardwareID NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @UID INT, @MaxDevices INT, @CurrentCount INT, @IsPremium BIT, @TrialExp DATETIME;

    -- 1. Identify User using Email and plain text Password
    SELECT @UID = UserID, @MaxDevices = AllowedDeviceLimit,
           @IsPremium = IsPremium, @TrialExp = TrialExpiryDate
    FROM Users
    WHERE Email = @InEmail
      AND [Password] = @InPassword
      AND IsActive = 1;

    -- 2. Error: User not found or Banned
    IF @UID IS NULL
    BEGIN
        SELECT 'INVALID_CREDENTIALS' AS Response; RETURN;
    END

    -- 3. Error: Trial check
    IF @IsPremium = 0 AND (@TrialExp IS NOT NULL AND @TrialExp < GETDATE())
    BEGIN
        SELECT 'TRIAL_EXPIRED' AS Response; RETURN;
    END

    -- 4. Success: Device is already registered
    IF EXISTS (SELECT 1 FROM UserDevices WHERE UserID = @UID AND HardwareID = @InHardwareID)
    BEGIN
        SELECT 'SUCCESS' AS Response, @IsPremium AS IsPremium, @UID AS UserID;
        UPDATE Users SET LastLogin = GETDATE() WHERE UserID = @UID;
        RETURN;
    END

    -- 5. New Device Check
    SELECT @CurrentCount = COUNT(*) FROM UserDevices WHERE UserID = @UID;

    IF @CurrentCount < @MaxDevices
    BEGIN
        INSERT INTO UserDevices (UserID, HardwareID) VALUES (@UID, @InHardwareID);
        SELECT 'SUCCESS' AS Response, @IsPremium AS IsPremium, @UID AS UserID;
    END
    ELSE
    BEGIN
        SELECT 'LIMIT_REACHED' AS Response, @UID AS UserID;
        SELECT DeviceID, HardwareID, AddedAt FROM UserDevices WHERE UserID = @UID;
    END
END;
GO
/****** Object:  StoredProcedure [dbo].[sp_RemoveDevice] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Table Name: Users, UserDevices
CREATE OR ALTER PROCEDURE [dbo].[sp_RemoveDevice]
    @InEmail NVARCHAR(100),
    @InPassword NVARCHAR(100),
    @InDeviceID INT
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @UID INT;

    -- 1. Verify credentials
    SELECT @UID = UserID
    FROM Users
    WHERE Email = @InEmail
      AND [Password] = @InPassword
      AND IsActive = 1;

    -- 2. Invalid credentials
    IF @UID IS NULL
    BEGIN
        SELECT 'INVALID_CREDENTIALS' AS Response; RETURN;
    END

    -- 3. Check device belongs to this user
    IF NOT EXISTS (SELECT 1 FROM UserDevices WHERE DeviceID = @InDeviceID AND UserID = @UID)
    BEGIN
        SELECT 'NOT_FOUND' AS Response; RETURN;
    END

    -- 4. Remove the device
    DELETE FROM UserDevices WHERE DeviceID = @InDeviceID AND UserID = @UID;

    SELECT 'SUCCESS' AS Response;
END;
GO
/****** Object:  Add FullName, Email, Mobile columns to Users ******/
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Users]') AND name = 'FullName')
    ALTER TABLE [dbo].[Users] ADD [FullName] NVARCHAR(100) NULL;
GO
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Users]') AND name = 'Email')
    ALTER TABLE [dbo].[Users] ADD [Email] NVARCHAR(100) NULL;
GO
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Users]') AND name = 'Mobile')
    ALTER TABLE [dbo].[Users] ADD [Mobile] NVARCHAR(20) NULL;
GO
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'UQ_Users_Email' AND object_id = OBJECT_ID(N'[dbo].[Users]'))
    ALTER TABLE [dbo].[Users] ADD CONSTRAINT UQ_Users_Email UNIQUE ([Email]);
GO
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'UQ_Users_Mobile' AND object_id = OBJECT_ID(N'[dbo].[Users]'))
    ALTER TABLE [dbo].[Users] ADD CONSTRAINT UQ_Users_Mobile UNIQUE ([Mobile]);
GO
/****** Object:  Table [dbo].[SubscriptionPlans] ******/
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SubscriptionPlans]') AND type = N'U')
BEGIN
CREATE TABLE [dbo].[SubscriptionPlans](
    [PlanID]         INT            IDENTITY(1,1) NOT NULL,
    [PlanName]       NVARCHAR(50)   NOT NULL,
    [AllowedModules] NVARCHAR(MAX)  NOT NULL,
    [IsActive]       BIT            NOT NULL DEFAULT (1),
PRIMARY KEY CLUSTERED ([PlanID] ASC),
CONSTRAINT [UQ_PlanName] UNIQUE ([PlanName])
) ON [PRIMARY]
END
GO
-- Seed default plans (only if table is empty)
IF NOT EXISTS (SELECT 1 FROM [dbo].[SubscriptionPlans])
BEGIN
    INSERT INTO [dbo].[SubscriptionPlans] (PlanName, AllowedModules) VALUES
    ('Trial',      '["GSTR2B","GSTR3B","GST_Challan"]'),
    ('Basic',      '["GSTR2B","GSTR3B","GSTR3B_Excel","GST_Verifier","GST_Challan","R1_JSON","JSON_Excel","R1_PDF","IMS","GSTR1_Cons","PDF_Merge","PDF_Split","PDF_Extract","PDF_Compress","PDF_Redact"]'),
    ('Pro',        '["GSTR2B","GSTR3B","GSTR3B_Excel","GST_Verifier","GST_Challan","R1_JSON","JSON_Excel","R1_PDF","IMS","GSTR1_Cons","IT_26AS","IT_Challan","ITR_Bot","Demand_Checker","Refund_Checker","PDF_Merge","PDF_Split","PDF_Extract","PDF_Compress","PDF_Redact"]'),
    ('Enterprise', '["GSTR2B","GSTR3B","GSTR3B_Excel","GST_Verifier","GST_Challan","R1_JSON","JSON_Excel","R1_PDF","IMS","GSTR1_Cons","IT_26AS","IT_Challan","ITR_Bot","Demand_Checker","Refund_Checker","PDF_Merge","PDF_Split","PDF_Extract","PDF_Compress","PDF_Redact","Bank_Excel","Email_GST_Request","Email_Invoice","Email_Payment","GST_Reco"]')
END
GO
-- Add PlanID column to Users if it doesn't exist
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[Users]') AND name = 'PlanID')
    ALTER TABLE [dbo].[Users] ADD [PlanID] INT NULL;
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE parent_object_id = OBJECT_ID(N'[dbo].[Users]') AND name = 'FK_Users_PlanID')
    ALTER TABLE [dbo].[Users] ADD CONSTRAINT FK_Users_PlanID FOREIGN KEY ([PlanID]) REFERENCES [dbo].[SubscriptionPlans] ([PlanID]);
GO
-- Back-fill PlanID for existing users (IsPremium=0 → Trial, IsPremium=1 → Enterprise)
UPDATE [dbo].[Users]
SET PlanID = (SELECT PlanID FROM [dbo].[SubscriptionPlans] WHERE PlanName = 'Trial')
WHERE PlanID IS NULL AND IsPremium = 0;
GO
UPDATE [dbo].[Users]
SET PlanID = (SELECT PlanID FROM [dbo].[SubscriptionPlans] WHERE PlanName = 'Enterprise')
WHERE PlanID IS NULL AND IsPremium = 1;
GO

/****** Object:  StoredProcedure [dbo].[sp_AuthenticateWithDeviceManagement] (updated) ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER PROCEDURE [dbo].[sp_AuthenticateWithDeviceManagement]
    @InEmail NVARCHAR(100),
    @InPassword NVARCHAR(100),
    @InHardwareID NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @UID INT, @MaxDevices INT, @CurrentCount INT, @IsPremium BIT, @TrialExp DATETIME;
    DECLARE @AllowedModules NVARCHAR(MAX), @PlanName NVARCHAR(50);

    -- 1. Identify User using Email and plain text Password
    SELECT @UID = UserID, @MaxDevices = AllowedDeviceLimit,
           @IsPremium = IsPremium, @TrialExp = TrialExpiryDate
    FROM Users
    WHERE Email = @InEmail
      AND [Password] = @InPassword
      AND IsActive = 1;

    -- 2. Error: User not found or Banned
    IF @UID IS NULL
    BEGIN
        SELECT 'INVALID_CREDENTIALS' AS Response; RETURN;
    END

    -- 3. Error: Trial check
    IF @IsPremium = 0 AND (@TrialExp IS NOT NULL AND @TrialExp < GETDATE())
    BEGIN
        SELECT 'TRIAL_EXPIRED' AS Response; RETURN;
    END

    -- 4. Fetch plan details
    SELECT @AllowedModules = sp.AllowedModules, @PlanName = sp.PlanName
    FROM SubscriptionPlans sp
    JOIN Users u ON u.PlanID = sp.PlanID
    WHERE u.UserID = @UID;

    -- 5. Success: Device is already registered
    IF EXISTS (SELECT 1 FROM UserDevices WHERE UserID = @UID AND HardwareID = @InHardwareID)
    BEGIN
        SELECT 'SUCCESS' AS Response, @IsPremium AS IsPremium, @UID AS UserID,
               @AllowedModules AS AllowedModules, @PlanName AS PlanName;
        UPDATE Users SET LastLogin = GETDATE() WHERE UserID = @UID;
        RETURN;
    END

    -- 6. New Device Check
    SELECT @CurrentCount = COUNT(*) FROM UserDevices WHERE UserID = @UID;

    IF @CurrentCount < @MaxDevices
    BEGIN
        INSERT INTO UserDevices (UserID, HardwareID) VALUES (@UID, @InHardwareID);
        SELECT 'SUCCESS' AS Response, @IsPremium AS IsPremium, @UID AS UserID,
               @AllowedModules AS AllowedModules, @PlanName AS PlanName;
    END
    ELSE
    BEGIN
        SELECT 'LIMIT_REACHED' AS Response, @UID AS UserID;
        SELECT DeviceID, HardwareID, AddedAt FROM UserDevices WHERE UserID = @UID;
    END
END;
GO

/****** Object:  StoredProcedure [dbo].[sp_RegisterTrialUser] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Table Name: Users
CREATE OR ALTER PROCEDURE [dbo].[sp_RegisterTrialUser]
    @InFullName  NVARCHAR(100),
    @InEmail     NVARCHAR(100),
    @InMobile    NVARCHAR(20),
    @InUsername  NVARCHAR(50),
    @InPassword  NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    -- 1. Uniqueness checks
    IF EXISTS (SELECT 1 FROM Users WHERE Email = @InEmail)
    BEGIN SELECT 'ERROR_EMAIL_EXISTS' AS Response; RETURN; END

    IF EXISTS (SELECT 1 FROM Users WHERE Mobile = @InMobile)
    BEGIN SELECT 'ERROR_MOBILE_EXISTS' AS Response; RETURN; END

    IF EXISTS (SELECT 1 FROM Users WHERE Username = @InUsername)
    BEGIN SELECT 'ERROR_USERNAME_EXISTS' AS Response; RETURN; END

    -- 2. Insert trial user (7-day trial, 1 device limit, Trial plan)
    DECLARE @TrialPlanID INT;
    SELECT @TrialPlanID = PlanID FROM SubscriptionPlans WHERE PlanName = 'Trial';

    INSERT INTO Users (FullName, Email, Mobile, Username, [Password], IsPremium, IsActive, AllowedDeviceLimit, TrialExpiryDate, PlanID)
    VALUES (@InFullName, @InEmail, @InMobile, @InUsername, @InPassword, 0, 1, 1, DATEADD(DAY, 7, GETDATE()), @TrialPlanID);

    SELECT 'SUCCESS' AS Response;
END;
GO
