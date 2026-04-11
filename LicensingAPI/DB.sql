/****** Object:  Table [dbo].[Admins]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Admins](
	[AdminID] [int] IDENTITY(1,1) NOT NULL,
	[Username] [nvarchar](50) NOT NULL,
	[Password] [nvarchar](100) NOT NULL,
	[FullName] [nvarchar](100) NULL,
	[Email] [nvarchar](100) NULL,
	[Role] [nvarchar](20) NOT NULL,
	[IsActive] [bit] NOT NULL,
	[CreatedAt] [datetime] NOT NULL,
	[LastLogin] [datetime] NULL,
	[CreatedBy] [int] NULL,
	[CanGrantPremium] [bit] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[AdminID] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UQ_Admins_Username] UNIQUE NONCLUSTERED 
(
	[Username] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SubscriptionPlans]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SubscriptionPlans](
	[PlanID] [int] IDENTITY(1,1) NOT NULL,
	[PlanName] [nvarchar](50) NOT NULL,
	[AllowedModules] [nvarchar](max) NOT NULL,
	[IsActive] [bit] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[PlanID] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UQ_PlanName] UNIQUE NONCLUSTERED 
(
	[PlanName] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[UserDevices]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
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
GO
/****** Object:  Table [dbo].[Users]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
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
	[FullName] [nvarchar](100) NULL,
	[Email] [nvarchar](100) NULL,
	[Mobile] [nvarchar](20) NULL,
	[PlanID] [int] NULL,
	[Profession] [nvarchar](100) NULL,
PRIMARY KEY CLUSTERED 
(
	[UserID] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[Username] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UQ_Users_Email] UNIQUE NONCLUSTERED 
(
	[Email] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UQ_Users_Mobile] UNIQUE NONCLUSTERED 
(
	[Mobile] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Admins] ADD  DEFAULT ('admin') FOR [Role]
GO
ALTER TABLE [dbo].[Admins] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[Admins] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[Admins] ADD  DEFAULT ((0)) FOR [CanGrantPremium]
GO
ALTER TABLE [dbo].[SubscriptionPlans] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[UserDevices] ADD  DEFAULT (getdate()) FOR [AddedAt]
GO
ALTER TABLE [dbo].[Users] ADD  DEFAULT ((0)) FOR [IsPremium]
GO
ALTER TABLE [dbo].[Users] ADD  DEFAULT ((1)) FOR [IsActive]
GO
ALTER TABLE [dbo].[Users] ADD  DEFAULT ((1)) FOR [AllowedDeviceLimit]
GO
ALTER TABLE [dbo].[Users] ADD  DEFAULT (getdate()) FOR [CreatedAt]
GO
ALTER TABLE [dbo].[Admins]  WITH CHECK ADD  CONSTRAINT [FK_Admins_CreatedBy] FOREIGN KEY([CreatedBy])
REFERENCES [dbo].[Admins] ([AdminID])
GO
ALTER TABLE [dbo].[Admins] CHECK CONSTRAINT [FK_Admins_CreatedBy]
GO
ALTER TABLE [dbo].[UserDevices]  WITH CHECK ADD FOREIGN KEY([UserID])
REFERENCES [dbo].[Users] ([UserID])
GO
ALTER TABLE [dbo].[UserDevices]  WITH CHECK ADD FOREIGN KEY([UserID])
REFERENCES [dbo].[Users] ([UserID])
GO
ALTER TABLE [dbo].[UserDevices]  WITH CHECK ADD FOREIGN KEY([UserID])
REFERENCES [dbo].[Users] ([UserID])
GO
ALTER TABLE [dbo].[UserDevices]  WITH CHECK ADD FOREIGN KEY([UserID])
REFERENCES [dbo].[Users] ([UserID])
GO
ALTER TABLE [dbo].[Users]  WITH CHECK ADD  CONSTRAINT [FK_Users_PlanID] FOREIGN KEY([PlanID])
REFERENCES [dbo].[SubscriptionPlans] ([PlanID])
GO
ALTER TABLE [dbo].[Users] CHECK CONSTRAINT [FK_Users_PlanID]
GO
/****** Object:  StoredProcedure [dbo].[sp_AuthenticateAdmin]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 CREATE PROCEDURE [dbo].[sp_AuthenticateAdmin]
      @Username NVARCHAR(50),
      @Password NVARCHAR(100)
  AS
  BEGIN
      IF EXISTS (SELECT 1 FROM Admins WHERE Username = @Username AND Password = @Password AND IsActive = 1)
          SELECT 'SUCCESS' AS Response, AdminID, Role, FullName, Email, CanGrantPremium
          FROM Admins
          WHERE Username = @Username AND Password = @Password AND IsActive = 1;
      ELSE
          SELECT 'INVALID_CREDENTIALS' AS Response, 0 AS AdminID, '' AS Role,
                 CAST('' AS NVARCHAR(100)) AS FullName,
                 CAST('' AS NVARCHAR(100)) AS Email,
                 CAST(0 AS BIT) AS CanGrantPremium;
  END
GO
/****** Object:  StoredProcedure [dbo].[sp_AuthenticateWithDeviceManagement]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_AuthenticateWithDeviceManagement]
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
    IF @IsPremium = 0 AND (@TrialExp IS NOT NULL AND @TrialExp < DATEADD(MINUTE, 330, GETUTCDATE()))
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
        UPDATE Users SET LastLogin = DATEADD(MINUTE, 330, GETUTCDATE()) WHERE UserID = @UID;
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
/****** Object:  StoredProcedure [dbo].[sp_CheckDeviceSession]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Checks if a device is already registered for the user.
-- Does NOT auto-register new devices. Used for silent session restore on app startup.
-- Returns SESSION_VALID if device is registered and account is active.
-- Returns SESSION_INVALID if device was removed (remote logout), account banned, or trial expired.
CREATE PROCEDURE [dbo].[sp_CheckDeviceSession]
    @InEmail      NVARCHAR(100),
    @InPassword   NVARCHAR(100),
    @InHardwareID NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @UID INT, @IsPremium BIT, @TrialExp DATETIME;
    DECLARE @AllowedModules NVARCHAR(MAX), @PlanName NVARCHAR(50);

    -- 1. Validate credentials and active status
    SELECT @UID = UserID, @IsPremium = IsPremium, @TrialExp = TrialExpiryDate
    FROM Users
    WHERE Email = @InEmail
      AND [Password] = @InPassword
      AND IsActive = 1;

    IF @UID IS NULL
    BEGIN
        SELECT 'SESSION_INVALID' AS Response; RETURN;
    END

    -- 2. Trial expiry check
    IF @IsPremium = 0 AND (@TrialExp IS NOT NULL AND @TrialExp < DATEADD(MINUTE, 330, GETUTCDATE()))
    BEGIN
        SELECT 'SESSION_INVALID' AS Response; RETURN;
    END

    -- 3. Device must already be registered (no auto-registration)
    IF NOT EXISTS (SELECT 1 FROM UserDevices WHERE UserID = @UID AND HardwareID = @InHardwareID)
    BEGIN
        SELECT 'SESSION_INVALID' AS Response; RETURN;
    END

    -- 4. Fetch plan details
    SELECT @AllowedModules = sp.AllowedModules, @PlanName = sp.PlanName
    FROM SubscriptionPlans sp
    JOIN Users u ON u.PlanID = sp.PlanID
    WHERE u.UserID = @UID;

    -- 5. Update last login and return success
    UPDATE Users SET LastLogin = DATEADD(MINUTE, 330, GETUTCDATE()) WHERE UserID = @UID;

    SELECT 'SESSION_VALID' AS Response, @IsPremium AS IsPremium, @UID AS UserID,
           @AllowedModules AS AllowedModules, @PlanName AS PlanName;
END;
GO
/****** Object:  StoredProcedure [dbo].[sp_CheckEmailExists]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Returns 1 if the email exists in Users, 0 otherwise.
CREATE PROCEDURE [dbo].[sp_CheckEmailExists]
    @InEmail NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT CAST(COUNT(1) AS BIT) AS EmailExists
    FROM   Users
    WHERE  Email = @InEmail;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_CreateAdmin]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
  CREATE PROCEDURE [dbo].[sp_CreateAdmin]
      @Username        NVARCHAR(50),
      @Password        NVARCHAR(100),
      @FullName        NVARCHAR(100) = NULL,
      @Email           NVARCHAR(100) = NULL,
      @Role            NVARCHAR(20)  = 'admin',
      @CreatedBy       INT           = NULL,
      @CanGrantPremium BIT           = 0
  AS
  BEGIN
      IF EXISTS (SELECT 1 FROM Admins WHERE Username = @Username)
      BEGIN
          SELECT 'USERNAME_TAKEN' AS Response; RETURN;
      END
      INSERT INTO Admins (Username, Password, FullName, Email, Role, IsActive, CreatedBy, CanGrantPremium, CreatedAt)
      VALUES (@Username, @Password, @FullName, @Email, @Role, 1, @CreatedBy, @CanGrantPremium, DATEADD(MINUTE, 330, GETUTCDATE()));
      SELECT 'SUCCESS' AS Response;
  END
GO
/****** Object:  StoredProcedure [dbo].[sp_DeleteAdmin]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_DeleteAdmin]
    @AdminID INT
AS
BEGIN
    SET NOCOUNT ON;

    -- Prevent deleting the last active superadmin
    IF (SELECT Role FROM Admins WHERE AdminID = @AdminID) = 'superadmin'
    BEGIN
        IF (SELECT COUNT(*) FROM Admins WHERE Role = 'superadmin' AND IsActive = 1) <= 1
        BEGIN SELECT 'ERROR_LAST_SUPERADMIN' AS Response; RETURN; END
    END

    -- Nullify CreatedBy references before deleting
    UPDATE Admins SET CreatedBy = NULL WHERE CreatedBy = @AdminID;

    DELETE FROM Admins WHERE AdminID = @AdminID;
    SELECT 'SUCCESS' AS Response;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_DeleteUserByAdmin]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_DeleteUserByAdmin]
    @UserID INT
AS
BEGIN
    SET NOCOUNT ON;

    -- Prevent orphan records by removing associated devices first
    DELETE FROM UserDevices WHERE UserID = @UserID;

    -- Now safe to delete the user
    DELETE FROM Users WHERE UserID = @UserID;

    SELECT 'SUCCESS' AS Response;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_GetAllAdmins]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 CREATE PROCEDURE [dbo].[sp_GetAllAdmins]
  AS
  BEGIN
      SELECT
          a.AdminID,
          a.Username,
          a.FullName,
          a.Email,
          a.Role,
          a.IsActive,
          a.CreatedAt,
          a.LastLogin,
          a.CreatedBy,
          c.Username   AS CreatedByUsername,
          a.Password,
          a.CanGrantPremium
      FROM Admins a
      LEFT JOIN Admins c ON a.CreatedBy = c.AdminID
      ORDER BY a.CreatedAt DESC;
  END
GO
/****** Object:  StoredProcedure [dbo].[sp_GetAllUsers]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 CREATE PROCEDURE [dbo].[sp_GetAllUsers]
  AS
  BEGIN
      SELECT
          u.UserID,
          u.Username,
          u.FullName,
          u.Email,
          u.Mobile,
          ISNULL(u.IsPremium, 0)           AS IsPremium,
          ISNULL(u.IsActive, 1)            AS IsActive,
          ISNULL(u.AllowedDeviceLimit, 1)  AS AllowedDeviceLimit,
          u.TrialExpiryDate,
          u.LastLogin,
          u.CreatedAt,
          sp.PlanName,
          (SELECT COUNT(*) FROM UserDevices d WHERE d.UserID = u.UserID) AS DeviceCount,
          u.Password
      FROM Users u
      LEFT JOIN SubscriptionPlans sp ON u.PlanID = sp.PlanID
      ORDER BY u.CreatedAt DESC;
  END
GO
/****** Object:  StoredProcedure [dbo].[sp_GetDashboardStats]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_GetDashboardStats]
AS
BEGIN
    SET NOCOUNT ON;

    SELECT 
        (SELECT COUNT(*) FROM Users) AS TotalUsers,
        (SELECT COUNT(*) FROM Users WHERE IsActive = 1) AS ActiveUsers,
        (SELECT COUNT(*) FROM Users WHERE IsPremium = 1) AS PremiumUsers,
        (SELECT COUNT(*) FROM Users WHERE IsPremium = 0) AS TrialUsers,
        (SELECT COUNT(*) FROM UserDevices) AS TotalDevicesRegistered;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_GrantUserPremium]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_GrantUserPremium]
      @UserID     INT,
      @PlanID     INT,
      @ExpiryDate DATETIME = NULL
  AS
  BEGIN
      SET NOCOUNT ON;

      UPDATE Users SET
          PlanID          = @PlanID,
          IsPremium       = 1,
          TrialExpiryDate = @ExpiryDate
      WHERE UserID = @UserID;

      SELECT
          u.Username,
          u.Email,
          u.Password,
          u.FullName,
          sp.PlanName
      FROM Users u
      LEFT JOIN SubscriptionPlans sp ON sp.PlanID = @PlanID
      WHERE u.UserID = @UserID;
  END;
GO
/****** Object:  StoredProcedure [dbo].[sp_RegisterTrialUser]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 CREATE PROCEDURE [dbo].[sp_RegisterTrialUser]
      @InFullName   NVARCHAR(100),
      @InEmail      NVARCHAR(100),
      @InMobile     NVARCHAR(20),
      @InUsername   NVARCHAR(50),
      @InPassword   NVARCHAR(100),
      @InProfession NVARCHAR(50) = NULL
  AS
  BEGIN
      SET NOCOUNT ON;

      IF EXISTS (SELECT 1 FROM Users WHERE Email = @InEmail)
      BEGIN SELECT 'ERROR_EMAIL_EXISTS' AS Response; RETURN; END

      IF EXISTS (SELECT 1 FROM Users WHERE Mobile = @InMobile)
      BEGIN SELECT 'ERROR_MOBILE_EXISTS' AS Response; RETURN; END

      IF EXISTS (SELECT 1 FROM Users WHERE Username = @InUsername)
      BEGIN SELECT 'ERROR_USERNAME_EXISTS' AS Response; RETURN; END

      DECLARE @TrialPlanID INT;
      SELECT @TrialPlanID = PlanID FROM SubscriptionPlans WHERE PlanName = 'Enterprise';

      INSERT INTO Users (FullName, Email, Mobile, Username, [Password], Profession, IsPremium, IsActive, AllowedDeviceLimit, TrialExpiryDate, PlanID)
      VALUES (@InFullName, @InEmail, @InMobile, @InUsername, @InPassword, @InProfession, 0, 1, 5, DATEADD(DAY, 3, DATEADD(MINUTE, 330, GETUTCDATE())), @TrialPlanID);

      SELECT 'SUCCESS' AS Response;
  END
GO
/****** Object:  StoredProcedure [dbo].[sp_RemoveDevice]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Table Name: Users, UserDevices
CREATE PROCEDURE [dbo].[sp_RemoveDevice]
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
/****** Object:  StoredProcedure [dbo].[sp_ToggleAdminStatus]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_ToggleAdminStatus]
    @AdminID INT
AS
BEGIN
    SET NOCOUNT ON;

    -- Prevent disabling the last active superadmin
    IF (SELECT Role FROM Admins WHERE AdminID = @AdminID) = 'superadmin'
       AND (SELECT IsActive FROM Admins WHERE AdminID = @AdminID) = 1
    BEGIN
        IF (SELECT COUNT(*) FROM Admins WHERE Role = 'superadmin' AND IsActive = 1) <= 1
        BEGIN SELECT 'ERROR_LAST_SUPERADMIN' AS Response; RETURN; END
    END

    UPDATE Admins SET IsActive = CASE WHEN IsActive = 1 THEN 0 ELSE 1 END WHERE AdminID = @AdminID;
    SELECT 'SUCCESS' AS Response;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ToggleUserPremium]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

CREATE PROCEDURE [dbo].[sp_ToggleUserPremium]
    @UserID INT
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @EnterprisePlanID INT;
    SELECT @EnterprisePlanID = PlanID FROM SubscriptionPlans WHERE PlanName = 'Enterprise';

    UPDATE Users SET
        IsPremium = CASE WHEN IsPremium = 1 THEN 0 ELSE 1 END,
        PlanID    = @EnterprisePlanID,
        TrialExpiryDate = CASE
            WHEN IsPremium = 1 THEN DATEADD(DAY, -1, DATEADD(MINUTE, 330, GETUTCDATE()))
            ELSE DATEADD(YEAR, 10, DATEADD(MINUTE, 330, GETUTCDATE()))
        END
    WHERE UserID = @UserID;

    SELECT 'SUCCESS' AS Response;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ToggleUserStatus]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_ToggleUserStatus]
    @UserID INT
AS
BEGIN
    SET NOCOUNT ON;

    UPDATE Users SET IsActive = CASE WHEN IsActive = 1 THEN 0 ELSE 1 END WHERE UserID = @UserID;
    SELECT 'SUCCESS' AS Response;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_UpdateAdmin]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 CREATE PROCEDURE [dbo].[sp_UpdateAdmin]
      @AdminID         INT,
      @FullName        NVARCHAR(100) = NULL,
      @Email           NVARCHAR(100) = NULL,
      @Role            NVARCHAR(20)  = 'admin',
      @IsActive        BIT           = 1,
      @Password        NVARCHAR(100) = NULL,
      @CanGrantPremium BIT           = 0
  AS
  BEGIN
      IF NOT EXISTS (SELECT 1 FROM Admins WHERE AdminID = @AdminID)
      BEGIN
          SELECT 'NOT_FOUND' AS Response; RETURN;
      END
      IF @Password IS NOT NULL AND @Password != ''
          UPDATE Admins SET
              FullName = @FullName, Email = @Email, Role = @Role,
              IsActive = @IsActive, Password = @Password, CanGrantPremium = @CanGrantPremium
          WHERE AdminID = @AdminID;
      ELSE
          UPDATE Admins SET
              FullName = @FullName, Email = @Email, Role = @Role,
              IsActive = @IsActive, CanGrantPremium = @CanGrantPremium
          WHERE AdminID = @AdminID;
      SELECT 'SUCCESS' AS Response;
  END
GO
/****** Object:  StoredProcedure [dbo].[sp_UpdatePassword]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Updates the password for the given email.
-- Returns 'SUCCESS' on update, 'NOT_FOUND' if no matching active account.
CREATE PROCEDURE [dbo].[sp_UpdatePassword]
    @InEmail    NVARCHAR(100),
    @InPassword NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (SELECT 1 FROM Users WHERE Email = @InEmail AND IsActive = 1)
    BEGIN
        SELECT 'NOT_FOUND' AS Response; RETURN;
    END

    UPDATE Users
    SET    [Password] = @InPassword
    WHERE  Email      = @InEmail
      AND  IsActive   = 1;

    SELECT 'SUCCESS' AS Response;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_WebLogin]    Script Date: 4/9/2026 5:48:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Used by the website to authenticate a user by email + password.
-- Does NOT check devices or plans — session only.
-- Returns user details on success, or empty result set on failure.
CREATE PROCEDURE [dbo].[sp_WebLogin]
    @InEmail    NVARCHAR(100),
    @InPassword NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT UserID, ISNULL(FullName, Email) AS FullName, Email
    FROM   Users
    WHERE  Email    = @InEmail
      AND  [Password] = @InPassword
      AND  IsActive = 1;

    -- Update last login if a row was found
    IF @@ROWCOUNT > 0
        UPDATE Users SET LastLogin = DATEADD(MINUTE, 330, GETUTCDATE()) WHERE Email = @InEmail;
END;

GO
