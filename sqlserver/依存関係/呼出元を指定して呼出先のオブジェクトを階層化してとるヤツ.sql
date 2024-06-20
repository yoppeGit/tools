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
		 0									-- 親の呼出レベル
		,OBJECT_NAME(referencing_id)		-- 親のオブジェクト名
		,referenced_schema_name				-- 子のオブジェクトが属するスキーマ
		,referenced_entity_name				-- 子のオブジェクト名
		,OBJECT_ID(referenced_entity_name)	-- 子のオブジェクトID
		,obj.type							-- 子のオブジェクトタイプ
		,obj.type_desc						-- 子のオブジェクト説明
    FROM
		sys.sql_expression_dependencies AS sed
		INNER JOIN sys.objects obj
			ON	obj.object_id = OBJECT_ID(referenced_entity_name)
    WHERE
		OBJECT_NAME(referencing_id) = @referencing_entity 
	UNION ALL
    SELECT
		 level + 1							-- 子の呼出レベル
		,OBJECT_NAME(sed.referencing_id)	-- 子のオブジェクト名
		,sed.referenced_schema_name			-- 孫のオブジェクトが属するスキーマ
		,sed.referenced_entity_name			-- 孫のオブジェクト名
		,OBJECT_ID(referenced_entity_name)	-- 孫のオブジェクトID
		,obj.type							-- 孫のオブジェクトタイプ
		,obj.type_desc						-- 孫のオブジェクト説明
    FROM
		-- 親テーブル
		ObjectDepends AS o
		-- 子テーブル１＿依存関係
		INNER JOIN sys.sql_expression_dependencies AS sed
			ON	sed.referencing_id = o.referenced_id
		-- 子テーブル２＿オブジェクト一覧
		INNER JOIN sys.objects obj
			ON	obj.object_id = sed.referenced_id
)
SELECT
	 level 呼出レベル
	,entity_name 呼出元オブジェクト名
	,referenced_schema 呼出先スキーマ名
	,referenced_entity 呼出先オブジェクト名
	,referenced_type 呼出先オブジェクトタイプ
FROM ObjectDepends
ORDER BY level
;

