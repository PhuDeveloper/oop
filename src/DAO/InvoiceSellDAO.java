/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import com.tpd.entity.InvoiceSell;
import com.tpd.helper.jdbcHelper;
import com.tpd.utils.XDate;
import java.util.ArrayList;
import java.util.List;
import java.sql.*;

/**
 *
 * @author ducit
 */
public class InvoiceSellDAO extends ShopDAO<InvoiceSell, Integer> {
    
    static String user = "sa";
    static String pass = "123456";
    static String url = "jdbc:sqlserver://localhost;databaseName=OOP";

    @Override
    public void insert(InvoiceSell e) {
        String sql = "INSERT INTO dbo.InvoiceSell\n"
                + "(idCustomer,idHumanSell,idVoucher,dateCreateInvoice,description,statusPay,statusInvoice, totalMoney, moneyCustom, moneyReturn)\n"
                + "VALUES(?, ?, ?, ?,?,?,?,?,?,?)";
        jdbcHelper.update(sql, e.getIdCustomer(), e.getIdHumanSell(), e.getIdVoucher(), e.getDateCreateInvoice(), e.getDescription(), e.getStatusPay(), e.getStatusInvoice(), e.getPrice(), e.getMoneyCustomer(), e.getMoneyReturn());
    }

    @Override
    public void update(InvoiceSell e) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void delete(Integer k) {
        String sql = "DELETE FROM dbo.InvoiceSELL where idInvoiceSell= ?";
      
        try {
            Connection con = DriverManager.getConnection(url, user, pass);
            PreparedStatement stmt = con.prepareStatement(sql);
            stmt.setString(1, Integer.toString(k));
            stmt.executeUpdate();
            con.close();
        } catch (SQLException ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        }
    }

    @Override
    public List<InvoiceSell> selectAll() {
        String sql = "SELECT * FROM dbo.InvoiceSell JOIN dbo.[User] ON [User].idUser = InvoiceSell.idHumanSell JOIN dbo.Customer ON Customer.idCustomer = InvoiceSell.idCustomer ORDER BY idInvoiceSell Desc";
        return selectBySql(sql);
    }

    @Override
    public InvoiceSell selectById(Integer k) {
        String sql = "SELECT * FROM dbo.InvoiceSell JOIN dbo.[User] ON [User].idUser = InvoiceSell.idHumanSell JOIN dbo.Customer ON Customer.idCustomer = InvoiceSell.idCustomer \n"
                + "where idInvoiceSell = ?";
        List<InvoiceSell> list = selectBySql(sql, k);
        if (list.isEmpty()) {
            return null;
        }
        return list.get(0);
    }

    @Override
    protected List<InvoiceSell> selectBySql(String sql, Object... args) {
        List<InvoiceSell> list = new ArrayList<>();
        try {
            ResultSet rs = jdbcHelper.query(sql, args);
            while (rs.next()) {
                InvoiceSell i = new InvoiceSell();
                i.setIdInvoiceSell(rs.getInt("idInvoiceSell"));
                i.setIdCustomer(rs.getInt("idCustomer"));
                i.setIdHumanSell(rs.getInt("idHumanSell"));
                i.setIdVoucher(rs.getInt("idVoucher"));
                i.setDateCreateInvoice(rs.getDate("dateCreateInvoice"));
                i.setDescription(rs.getString("description"));
                i.setStatusPay(rs.getBoolean("statusPay"));
                i.setStatusInvoice(rs.getBoolean("statusInvoice"));
                i.setNameCustomer(rs.getString(23));
                i.setNameUser(rs.getString(13));
                i.setPrice(rs.getDouble("totalMoney"));
                i.setMoneyCustomer(rs.getDouble("moneyCustom"));
                i.setMoneyReturn(rs.getDouble("moneyReturn"));
                list.add(i);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    public Float getTotalMoney(Integer idInvoice) {
        String sql = "SELECT idInvoiceSell, SUM(detailsInvoiceSELL.quatity * price)\n"
                + "AS N'Total'\n"
                + "FROM dbo.detailsInvoiceSELL\n"
                + "GROUP BY idInvoiceSell\n"
                + "HAVING idInvoiceSell = ?";
        try {
            ResultSet rs = jdbcHelper.query(sql, idInvoice);
            while (rs.next()) {
                return rs.getFloat("Total");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public int totalPage(String Stringdate) {
        ResultSet rs;
        if (!Stringdate.isEmpty()) {
            java.util.Date date = XDate.toDate(Stringdate, "yyyy-MM-dd");
            String sql = " select count(*) as soLuong from InvoiceSell WHERE  dateCreateInvoice BETWEEN '" + new java.sql.Date(date.getTime()) + " 00:00:00.000'" + "AND '" + new java.sql.Date(date.getTime()) + " 23:59:59.000' ";
            try {
                rs = jdbcHelper.query(sql);
                while (rs.next()) {
                    return rs.getInt("soLuong");
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        String sql = "select count(*) as soLuong from InvoiceSell";
        try {
            rs = jdbcHelper.query(sql);
            while (rs.next()) {
                return rs.getInt("soLuong");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return 0;
    }

    public List<InvoiceSell> pagingPage(int page, int pageSize, String Stringdate) {
        if (!Stringdate.isEmpty()) {
            java.util.Date date = XDate.toDate(Stringdate, "yyyy-MM-dd");
            String sql = " SELECT * FROM dbo.InvoiceSell JOIN dbo.[User] ON [User].idUser = InvoiceSell.idHumanSell JOIN dbo.Customer ON Customer.idCustomer = InvoiceSell.idCustomer\n"
                    + "WHERE InvoiceSell.statusInvoice <> 0 and dateCreateInvoice BETWEEN '" + new java.sql.Date(date.getTime()) + " 00:00:00.000'" + "AND '" + new java.sql.Date(date.getTime()) + " 23:59:59.000' \n"
                    + "order by idInvoiceSell desc OFFSET ? ROWS FETCH FIRST ? ROWS ONLY";
            return selectBySql(sql, (page - 1) * pageSize, pageSize);
        }
        String sql = "SELECT * FROM dbo.InvoiceSell JOIN dbo.[User] ON [User].idUser = InvoiceSell.idHumanSell JOIN dbo.Customer ON Customer.idCustomer = InvoiceSell.idCustomer \n"
                + "order by idInvoiceSell desc OFFSET ? ROWS FETCH FIRST ? ROWS ONLY";
        return selectBySql(sql, (page - 1) * pageSize, pageSize);
    }

}
