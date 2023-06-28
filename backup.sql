select
    "VRoleID",
    max("LovePointLevs") "LovePointLevs",
    max("TotalOnlineTime") "TotalOnlineTime",
    max("Level") "Level",
    max("RegisterDate") "RegisterDate"
from
    v_event_68
where
    (
        "$part_event" IN ('PlayerLogout')
        and "$part_date" = '2023-03-18'
    )
group by
    "VRoleID"
order by
    max("DtEventTime") desc
    /*
     --------------------------------------------------
     */
select
    "VRoleID",
    max("LovePointLevs") "LovePointLevs",
    max("TotalOnlineTime") "TotalOnlineTime",
    max("Level") "Level",
    max("RegisterDate") "RegisterDate"
from
    v_event_68
where
    (
        "$part_event" IN ('PlayerLogout')
        and "$part_date" = '2023-03-18'
        and RegisterDate < timestamp '2023-03-17 23:55:00'
    )
group by
    "VRoleID"
order by
    max("DtEventTime") desc
    /*
     --------------------------------------------------
     */
select
    DISTINCT "Reason",
    "SubReason",
    "AddOrReduce",
    "TargetParam",
    "IMoneyType",
    "IMoney"
from
    v_event_68
where
    (
        "$part_event" IN ('MoneyFlow')
        and "$part_date" BETWEEN '2023-03-17'
        and '2023-03-23'
        and RegisterDate < timestamp '2023-03-17 23:55:00'
        and "Reason" = 102
    )
order by
    "Reason" desc