USE [QualiAdmFin]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


/*
------------------------------------------------------------------------------------------------------
	DATA:		13/03/2022
	AUTOR: 		Paulo Sérgio Garcia
	EXEMPLO: 	EXECUTE sp_compila_objetos
------------------------------------------------------------------------------------------------------
*/

CREATE OR ALTER PROCEDURE [dbo].[sp_compila_objetos]
AS
BEGIN

	SET NOCOUNT ON

	DECLARE @object_name		NVARCHAR(128),
			@object_definition	NVARCHAR(MAX),
			@object_type		CHAR(2),
			@message			NVARCHAR(MAX),
			@error_message 		NVARCHAR(4000),
			@error_number		INT,
			@error_severity 	INT,
			@error_state		INT
			
	DECLARE object_cursor CURSOR FOR
		SELECT objects.name,
			   objects.type,
			   sql_modules.definition
		  FROM sys.sql_modules
			   JOIN
			   (SELECT name,
					   object_id,
					   type
				  FROM sys.all_objects 
				 WHERE schema_id in (SELECT schema_id FROM sys.schemas WHERE name ='dbo')
				   AND is_ms_shipped <> 1										-- Objeto não embarcado por Microsoft
				   AND object_id > 0											-- Objeto não é do sistema
				   AND (type = 'P' OR type = 'V')								-- Objeto é uma Procedure ou View
				) objects ON sql_modules.object_id = objects.object_id
		  ORDER BY objects.type, objects.name

	OPEN object_cursor

	FETCH NEXT FROM object_cursor INTO @object_name, @object_type, @object_definition			-- Le a 1ª linha do cursor

	WHILE @@FETCH_STATUS = 0
	BEGIN

		SET @message = N'Processando objeto tipo: ' + @object_type + ', nome: ' + @object_name + ' ... '
		RAISERROR(@message, 0, 1) WITH NOWAIT
		
		BEGIN TRY 
		
			IF (@object_type = 'P')						-- Recompila procedures
				BEGIN
					EXECUTE sp_recompile @object_name 			
				END
			ELSE										-- Drop / Create a view
				BEGIN
				    EXEC('DROP VIEW ' + @object_name)
					EXEC(@object_definition)		
				END

		END TRY  

		BEGIN CATCH
		
			SELECT @error_number   = ERROR_NUMBER(),
				   @error_message  = ERROR_MESSAGE(),  
				   @error_severity = ERROR_SEVERITY(),  
				   @error_state    = ERROR_STATE()  
						
			SET @message = N'O objeto ' + @object_name + ' apresentou erro em sua execução (' + STR(@error_number) + ' / ' + @error_message + ')'
			RAISERROR(@message, @error_severity, @error_state) WITH NOWAIT
				   
		END CATCH 

		FETCH NEXT FROM object_cursor INTO @object_name, @object_type, @object_definition		-- Le as próximas linhas do cursor

	END
	
	CLOSE object_cursor
	
	DEALLOCATE object_cursor
	
END

GO


