DECLARE @referencing_entity AS sysname = N'mainProc';

WITH ObjectDepends(
	 level
	,entity_name
	,referenced_schema
	,referenced_entity
	,referenced_id
	,referenced_type
	,referenced_type_desc
)
AS(
    SELECT
		 0									-- �e�̌ďo���x��
		,OBJECT_NAME(referencing_id)		-- �e�̃I�u�W�F�N�g��
		,referenced_schema_name				-- �q�̃I�u�W�F�N�g��������X�L�[�}
		,referenced_entity_name				-- �q�̃I�u�W�F�N�g��
		,OBJECT_ID(referenced_entity_name)	-- �q�̃I�u�W�F�N�gID
		,obj.type							-- �q�̃I�u�W�F�N�g�^�C�v
		,obj.type_desc						-- �q�̃I�u�W�F�N�g����
    FROM
		sys.sql_expression_dependencies AS sed
		INNER JOIN sys.objects obj
			ON	obj.object_id = OBJECT_ID(referenced_entity_name)
    WHERE
		OBJECT_NAME(referencing_id) = @referencing_entity 
	UNION ALL
    SELECT
		 level + 1							-- �q�̌ďo���x��
		,OBJECT_NAME(sed.referencing_id)	-- �q�̃I�u�W�F�N�g��
		,sed.referenced_schema_name			-- ���̃I�u�W�F�N�g��������X�L�[�}
		,sed.referenced_entity_name			-- ���̃I�u�W�F�N�g��
		,OBJECT_ID(referenced_entity_name)	-- ���̃I�u�W�F�N�gID
		,obj.type							-- ���̃I�u�W�F�N�g�^�C�v
		,obj.type_desc						-- ���̃I�u�W�F�N�g����
    FROM
		-- �e�e�[�u��
		ObjectDepends AS o
		-- �q�e�[�u���P�Q�ˑ��֌W
		INNER JOIN sys.sql_expression_dependencies AS sed
			ON	sed.referencing_id = o.referenced_id
		-- �q�e�[�u���Q�Q�I�u�W�F�N�g�ꗗ
		INNER JOIN sys.objects obj
			ON	obj.object_id = sed.referenced_id
)
SELECT
	 level �ďo���x��
	,entity_name �ďo���I�u�W�F�N�g��
	,referenced_schema �ďo��X�L�[�}��
	,referenced_entity �ďo��I�u�W�F�N�g��
	,referenced_type �ďo��I�u�W�F�N�g�^�C�v
FROM ObjectDepends
ORDER BY level
;

