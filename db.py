import psycopg2
from dotenv import dotenv_values



config = dotenv_values(".env") 
class ConnexionOdoo:

    def connect():
        conn_string = "host={0} dbname={1} user={2} password={3}".format(config["DB_HOST"],config["DB_NAME"],config["DB_USER"],config["DB_PASSWORD"])
        print ("Connecting to database\n	->%s" % (conn_string))
        conn = psycopg2.connect(conn_string)
        cursor = conn.cursor()
        return cursor

    def getEmployeeContracts(self,company_id):
        cursor = self.connect()
        cursor.execute("""SELECT
                            hc.id,concat(he.firstname ,' ',he.name_related) "Employee", 
                            hc.date_start, 
                            hc.date_end, 
                            concat(he2.firstname ,' ',he2.name_related) as "responsable name" , 
                            he2.mobile_phone  as "responsable Phone number" ,
                            he2.work_phone  as "responsable Phone number 2" 
                            from hr_contract hc 
                            right join hr_employee he on hc.employee_id = he.id
                            right join hr_employee he2 on hc.parent_id = he2.id
                            where 
                                hc.date_end > current_date 
                                and extract(month from hc.date_end)= extract(month from DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 MONTH')
                                and extract(year from hc.date_end)>= extract(year from CURRENT_DATE) 
                                and extract(year from hc.date_end)<= extract(year from CURRENT_DATE) +1
                                and  hc.company_id =%s
                            order by hc.date_end asc
                         """
                        %company_id
                        )
        records = cursor.fetchall()
        cursor.close()
        return records