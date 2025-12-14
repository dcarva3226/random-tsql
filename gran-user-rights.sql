-----------------------------------------------------------------------
-- Script: GrantUserSurveyRights.sql
-- Purpose: Grant a user rights to Survey.
-- Author: D.C.
-- Date: 10.06.2004
-----------------------------------------------------------------------
USE SSISurvey

DECLARE @CommAdmin int
DECLARE @CollAdmin int
DECLARE @DomainId int
DECLARE @Domain varchar(100)
DECLARE @LogonName varchar(100)
DECLARE @SubCollAdmin int
DECLARE @UserId int
DECLARE @UserName varchar(100)

-- User-specified values ---------------------

SET @Domain 	= 'DOM'
SET @LogonName = 'DOM\Administrator'
SET @UserName 	= 'Administrator'
SET @CommAdmin = 2224184
SET @CollAdmin = 2510720
SET @SubCollAdmin = 19840
SET @UserID    = ( SELECT UserID FROM SSIUser WHERE LogonName = @LogonName ) -- in case it already exists
-- End user-specified values -----------------

-- Insert SSIDomain record unless it already exists
IF NOT EXISTS( SELECT DomainId FROM SSIDomain WHERE Name = @Domain )
  BEGIN
	INSERT INTO SSIDomain ( Name, DomainType, Status ) 
	VALUES ( @Domain, 4, 1 )
	SET @DomainId = ( SELECT @@IDENTITY )
  END
ELSE
  BEGIN
	SET @DomainId = ( SELECT DomainId FROM SSIDomain WHERE Name = @Domain )
  END

-- Insert SSIUser record unless it already exists
IF ISNULL(@UserID, -999) = -999
  BEGIN
	INSERT INTO SSIUser ( LogonName, UserName, DomainId, Status, InPopulation, UserType ) 
	VALUES ( @LogonName, @UserName, @DomainId, 1, 1, 1 )
	SET @UserId = ( SELECT @@IDENTITY )
  END
ELSE
  BEGIN
     -- Delete the existing users membership rules so we don't get dupes
     DELETE FROM CommunityMembership WHERE ResourceID = @UserID
     DELETE FROM CommunityRule WHERE [Value] = @LogonName
  END

--Grant rights in CommunityRule
DECLARE @CommId int
DECLARE @ResType int
DECLARE cur_cursor CURSOR FOR
SELECT CommunityId, ResourceType
FROM Community

OPEN cur_cursor

FETCH NEXT FROM cur_cursor
INTO @CommId, @ResType

IF @@FETCH_STATUS <> 0
	BEGIN
		PRINT '0'
	END
ELSE
	BEGIN
		WHILE @@FETCH_STATUS = 0
		BEGIN
			INSERT INTO CommunityRule ( CommunityId, RuleType, PropertyClass, Contributor, PropertyName, Powers, Operator, Value )
			VALUES( @CommId, 8, 4, 0, 'LogonName', CASE WHEN @ResType = 2 AND @CommId = 2 THEN @CollAdmin WHEN @ResType = 2 AND @CommId <> 2 THEN @SubCollAdmin ELSE @CommAdmin END, '=', @LogonName )

			INSERT INTO CommunityMembership ( CommunityId, ResourceId, ResourceType, Contributor, Powers )
			VALUES ( @CommId, @UserId, 1, 0, CASE WHEN @ResType = 2 THEN @CollAdmin ELSE @CommAdmin END )

			FETCH NEXT FROM cur_cursor
			INTO @CommId, @ResType
		END
	END

CLOSE cur_cursor
DEALLOCATE cur_cursor

