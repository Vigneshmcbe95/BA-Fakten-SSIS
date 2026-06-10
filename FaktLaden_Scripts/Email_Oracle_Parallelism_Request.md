# Email Draft — Central Per-User Parallelism for PolyBase Reads (Oracle 19c)

**Subject:** Request: Central per-user parallelism setup for PolyBase reads (Oracle 19c)

Hi team,

We read several large fact tables from Oracle (19c) into SQL Server via **PolyBase external tables**.

**The problem:** We cannot set an Oracle parallel hint (`/*+ PARALLEL */`) from the SQL Server side. PolyBase generates the remote query itself, and there is no way to inject an Oracle optimizer hint — neither in the `SELECT` against the external table, nor via the external data source / connection options (those carry connection settings only, not query hints). So parallelism has to be controlled **on the Oracle side**.

**What we'd like:** Rather than setting `PARALLEL` on each individual table, we'd prefer this handled **centrally for the PolyBase service user**. Two options we believe fit:

1. **Resource Manager (cap):** map our PolyBase Oracle user (`<USERNAME>`) to a consumer group with a `PARALLEL_DEGREE_LIMIT_P1` directive (`DBMS_RESOURCE_MANAGER.CREATE_PLAN_DIRECTIVE`) — capping max DOP per query for that user.
2. **Logon trigger (enable):** an `AFTER LOGON` trigger for that user running `ALTER SESSION FORCE PARALLEL QUERY PARALLEL <n>` (or `parallel_degree_policy = AUTO`) so its reads actually go parallel.

Could you advise the preferred central approach on our 19c instance, and a suitable DOP value? Happy to discuss.

Thanks,
Vignesh

---

## Reference documentation (Oracle 19c)

- **DBMS_RESOURCE_MANAGER — `CREATE_PLAN_DIRECTIVE` (section 150.4.11), parameter `parallel_degree_limit_p1`:**
  https://docs.oracle.com/en/database/oracle/oracle-database/19/arpls/DBMS_RESOURCE_MANAGER.html#GUID-8EC6C735-338D-46D4-B346-AD16D0622B30
- **Specifying a Degree of Parallelism Limit for Consumer Groups (VLDB & Partitioning Guide):**
  https://docs.oracle.com/database/121/VLDBG/GUID-44222E39-87C7-4206-8FE3-91DAF0D3573C.htm
- **Managing Resources with Oracle Database Resource Manager — session-to-consumer-group mapping by `ORACLE_USER` (Admin Guide):**
  https://docs.oracle.com/en/database/oracle/oracle-database/19/admin/managing-resources-with-oracle-database-resource-manager.html
- **DBMS_RESOURCE_MANAGER_PRIVS — `GRANT_SWITCH_CONSUMER_GROUP` (required for the mapping to take effect):**
  https://docs.oracle.com/cd/E18283_01/appdev.112/e16760/d_resmpr.htm
