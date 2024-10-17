drop table if exists #TMP_CHOFUKU
create table #TMP_CHOFUKU(
	 出力順 decimal(10)
	,開始日 varchar(10)
	,終了日 varchar(10)
	,メモ varchar(200)
)

--------------------------------
-- CASE1
--------------------------------
insert into #TMP_CHOFUKU values
 (null, '2024/10/05', '2024/12/31', null)
,(null, '2024/10/01', '2024/10/05', null)
,(null, '2024/08/19', '2024/10/02', null)
,(null, '2025/01/06', '2025/02/03', null)
,(null, '2025/01/06', '2025/02/03', null)

update TMP1
set
	TMP1.出力順 = TMP2.出力順
from
	#TMP_CHOFUKU TMP1
	inner join (
		select
			 開始日
			,終了日
			,ROW_NUMBER() over (order by 開始日 asc, 終了日 asc) 出力順
		from
			#TMP_CHOFUKU
	) TMP2
	on	TMP1.開始日 = TMP2.開始日
	and	tmp1.終了日 = tmp2.終了日

select * from #TMP_CHOFUKU order by 出力順 asc

select
	*
from
	#TMP_CHOFUKU TBL1
	inner join #TMP_CHOFUKU TBL2
		on	TBL1.開始日 <= TBL2.終了日
		and	TBL2.開始日 <= TBL1.終了日
		and	TBL1.出力順 < TBL2.出力順
order by
	TBL1.出力順 asc, TBL2.出力順 asc

--------------------------------
-- CASE2
--------------------------------
truncate table #TMP_CHOFUKU
insert into #TMP_CHOFUKU values
 (null, '2024/10/01', '2024/10/05', null)
,(null, '2024/08/19', '2024/09/30', null)
,(null, '2025/01/06', '2025/02/03', null)

update TMP1
set
	TMP1.出力順 = TMP2.出力順
from
	#TMP_CHOFUKU TMP1
	inner join (
		select
			 開始日
			,終了日
			,ROW_NUMBER() over (order by 開始日 asc, 終了日 asc) 出力順
		from
			#TMP_CHOFUKU
	) TMP2
	on	TMP1.開始日 = TMP2.開始日
	and	tmp1.終了日 = tmp2.終了日

select * from #TMP_CHOFUKU order by 出力順 asc

select
	*
from
	#TMP_CHOFUKU TBL1
	inner join #TMP_CHOFUKU TBL2
		on	TBL1.開始日 <= TBL2.終了日
		and	TBL2.開始日 <= TBL1.終了日
		and	TBL1.出力順 < TBL2.出力順
order by
	TBL1.出力順 asc, TBL2.出力順 asc

drop table if exists #TMP_CHOFUKU
