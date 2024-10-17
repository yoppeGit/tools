drop table if exists #TMP_CHOFUKU
create table #TMP_CHOFUKU(
	 �o�͏� decimal(10)
	,�J�n�� varchar(10)
	,�I���� varchar(10)
	,���� varchar(200)
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
	TMP1.�o�͏� = TMP2.�o�͏�
from
	#TMP_CHOFUKU TMP1
	inner join (
		select
			 �J�n��
			,�I����
			,ROW_NUMBER() over (order by �J�n�� asc, �I���� asc) �o�͏�
		from
			#TMP_CHOFUKU
	) TMP2
	on	TMP1.�J�n�� = TMP2.�J�n��
	and	tmp1.�I���� = tmp2.�I����

select * from #TMP_CHOFUKU order by �o�͏� asc

select
	*
from
	#TMP_CHOFUKU TBL1
	inner join #TMP_CHOFUKU TBL2
		on	TBL1.�J�n�� <= TBL2.�I����
		and	TBL2.�J�n�� <= TBL1.�I����
		and	TBL1.�o�͏� < TBL2.�o�͏�
order by
	TBL1.�o�͏� asc, TBL2.�o�͏� asc

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
	TMP1.�o�͏� = TMP2.�o�͏�
from
	#TMP_CHOFUKU TMP1
	inner join (
		select
			 �J�n��
			,�I����
			,ROW_NUMBER() over (order by �J�n�� asc, �I���� asc) �o�͏�
		from
			#TMP_CHOFUKU
	) TMP2
	on	TMP1.�J�n�� = TMP2.�J�n��
	and	tmp1.�I���� = tmp2.�I����

select * from #TMP_CHOFUKU order by �o�͏� asc

select
	*
from
	#TMP_CHOFUKU TBL1
	inner join #TMP_CHOFUKU TBL2
		on	TBL1.�J�n�� <= TBL2.�I����
		and	TBL2.�J�n�� <= TBL1.�I����
		and	TBL1.�o�͏� < TBL2.�o�͏�
order by
	TBL1.�o�͏� asc, TBL2.�o�͏� asc

drop table if exists #TMP_CHOFUKU
