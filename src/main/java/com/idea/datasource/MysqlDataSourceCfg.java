package com.idea.datasource;

import com.alibaba.druid.pool.DruidDataSource;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.springframework.context.annotation.Configuration;
import org.springframework.stereotype.Repository;

import javax.sql.DataSource;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

/**
 * @ClassName: MysqlDataSourceCfg
 * @Description: Mysql的数据源配置，未启用，改用springboot的jdbcTemplate方式 <br>
 * @Author: yuanbao
 * @Date: 2025/3/10
 **/
@Repository
@Configuration
public class MysqlDataSourceCfg {
    private static final Logger logger = LoggerFactory.getLogger(MysqlDataSourceCfg.class);

    private static final DruidDataSource ds; //aps中间库
    private static final DruidDataSource ds2; //人员同步接口


    static {
        ds = new DruidDataSource();// DruidDataSourceFactory.createDataSource(pp);
        ds.setName("apsdb");//
        ds.setUrl("jdbc:mysql://10.1.2.100:3306/unimax55_develop?characterEncoding=utf8&useSSL=false");
        ds.setDriverClassName("com.mysql.jdbc.Driver");
        ds.setUsername("root");
        ds.setPassword("123456");
        ds.setInitialSize(50);
        ds.setMinIdle(10);
        ds.setMaxActive(500);
        ds.setMaxWait(60000);
        ds.setMinEvictableIdleTimeMillis(300000);
        ds.setRemoveAbandoned(true);// 超过时间限制是否回收
        ds.setRemoveAbandonedTimeout(60000);// 超过时间限制多长
        ds.setTimeBetweenEvictionRunsMillis(60000);// 配置间隔多久才进行一次检测，检测需要关闭的空闲连接，单位是毫秒
        ds.setTestWhileIdle(true);
        ds.setTestOnBorrow(false);
        ds.setTestOnReturn(false);
        ds.setPoolPreparedStatements(true);
        ds.setMaxPoolPreparedStatementPerConnectionSize(20);
        ds.setValidationQuery("SELECT 1");
        ds.setTimeBetweenLogStatsMillis(3600000);
        ds.addConnectionProperty("remarksReporting", "true");

    }
    static
    {
        ds2 = new DruidDataSource();// DruidDataSourceFactory.createDataSource(pp);

        ds2.setName("hzPlatform");
        ds2.setUrl("jdbc:sqlserver://10.10.66.101:1433;DatabaseName=hzPlatform");
        ds2.setDriverClassName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
        ds2.setUsername("sa");
        ds2.setPassword("syecrg@7503");
        ds2.setInitialSize(50);
        ds2.setMinIdle(10);
        ds2.setMaxActive(500);
        ds2.setMaxWait(60000);
        ds2.setMinEvictableIdleTimeMillis(300000);
        ds2.setRemoveAbandoned(true);// 超过时间限制是否回收
        ds2.setRemoveAbandonedTimeout(180);// 超过时间限制多长
        ds2.setTimeBetweenEvictionRunsMillis(60000);// 配置间隔多久才进行一次检测，检测需要关闭的空闲连接，单位是毫秒
        ds2.setTestWhileIdle(true);
        ds2.setTestOnBorrow(false);
        ds2.setTestOnReturn(false);
        ds2.setPoolPreparedStatements(true);
        ds2.setMaxPoolPreparedStatementPerConnectionSize(20);
        ds2.setValidationQuery("SELECT 1");
        ds2.setTimeBetweenLogStatsMillis(3600000);
        ds2.addConnectionProperty("remarksReporting", "true");

    }

    public static DataSource getDataSource() {
        return ds;
    }

    public static Connection getConnection() {
        try {
            return ds.getConnection();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }
    public static DataSource getDataSource2()
    {
        return ds2;
    }
    public static Connection getConnection2()
    {
        try
        {
            return ds2.getConnection();
        } catch (SQLException e)
        {
            throw new RuntimeException(e);
            //throw new MestarException(e,"数据库连接失败");
        }
    }
    public static void close(Connection conn, Statement stmt, ResultSet rs) {
        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException e) {
                logger.error("Error closing ResultSet", e);
            }
        }
        if (stmt != null) {
            try {
                stmt.close();
            } catch (SQLException e) {
                logger.error("Error closing Statement", e);
            }
        }
        if (conn != null) {
            try {
                conn.close();
            } catch (SQLException e) {
                logger.error("Error closing Connection", e);
            }
        }
    }

    public static void close(Connection conn, Statement stmt) {
        close(conn, stmt, null);
    }
}
