---------------------------------------
--テーブル定義を取得
---------------------------------------
declare @tblName nvarchar(200) = '%%'		--引数.テーブル名

SELECT
	 cols.column_id	as [No.]						--通し番号
    ,tbls.name AS テーブル名							--テーブル名
    ,cols.name AS カラム名							--カラム名
    ,type_name(cols.user_type_id) AS データ型			--データ型
	,case											--データ桁
		when type_name(cols.user_type_id) = 'char'
			then convert(varchar, cols.max_length)
		when type_name(cols.user_type_id) like '%char%'
			then convert(varchar, cols.max_length/2)
		when type_name(cols.user_type_id) = 'decimal'
			then concat(cols.precision, ',', cols.scale)
		else
			''
	 end as 桁・精度
	,case											--主キー有無
		when idx_cols.column_id is not null then 'PK'
		else ''
	 end as PK
	,CASE											--NULL許可
        WHEN cols.is_nullable = 1 THEN '○'
        ELSE '×'
     END AS [NULL]
	,CASE											--identify有無
		WHEN cols.is_identity = 1
			THEN concat(
					 'identity(', convert(varchar, iden_cols.seed_value)
					,',', convert(varchar, iden_cols.increment_value), ')'
					)
		ELSE
			''
	 END AS [IDENTITY]
FROM
    sys.tables AS tbls
    
	inner JOIN sys.columns AS cols ON
        tbls.object_id = cols.object_id
    
	left outer JOIN sys.key_constraints AS key_const ON
        tbls.object_id = key_const.parent_object_id
		AND key_const.type = 'PK'
    
	left outer join sys.index_columns AS idx_cols ON
        key_const.parent_object_id = idx_cols.object_id
        AND key_const.unique_index_id  = idx_cols.index_id
		AND tbls.object_id = idx_cols.object_id
		AND cols.column_id = idx_cols.index_column_id
	
	LEFT OUTER JOIN sys.identity_columns as iden_cols on
		cols.object_id = iden_cols.object_id
		AND cols.column_id = iden_cols.column_id

where 1=1
	and tbls.name like @tblName

order by
	tbls.name asc, cols.column_id asc

