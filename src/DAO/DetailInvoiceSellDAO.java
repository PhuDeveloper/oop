/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import com.tpd.entity.DetailInvoiceSell;
import com.tpd.helper.jdbcHelper;
import java.util.ArrayList;
import java.util.List;
import java.sql.*;

/**
 *
 * @author ducit
 */
public class DetailInvoiceSellDAO extends ShopDAO<DetailInvoiceSell, Integer> {
    static String user = "sa";
    static String pass = "123456";
    static String url = "jdbc:sqlserver://localhost;databaseName=OOP";

    @Override
    public void insert(DetailInvoiceSell e) {
        String sql = "INSERT INTO dbo.detailsInvoiceSELL\n"
                + "(idInvoiceSell,idPrDetails,quatity,price)\n"
                + "VALUES\n"
                + "((SELECT TOP 1  idInvoiceSell FROM dbo.InvoiceSell ORDER BY idInvoiceSell DESC),?, ?, ?)";
        jdbcHelper.update(sql, e.getIdPrDetails(), e.getQuantity(), e.getPrice());
    }

    @Override
    public void update(DetailInvoiceSell e) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void delete(Integer k) {
        String sql = "DELETE FROM dbo.detailsInvoiceSELL where idInvoiceSell= ?";
      
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
    public List<DetailInvoiceSell> selectAll() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public DetailInvoiceSell selectById(Integer k) {
        String sql = "SELECT * FROM dbo.detailsInvoiceSELL where idDetailsInvoiceSELL= ?";
        DetailInvoiceSell result = selectBySql(sql).get(0);
        return result;
    }

    @Override
    protected List<DetailInvoiceSell> selectBySql(String sql, Object... args) {
        List<DetailInvoiceSell> list = new ArrayList<>();
        try {
            ResultSet rs = jdbcHelper.query(sql, args);
            while (rs.next()) {
                DetailInvoiceSell de = new DetailInvoiceSell();
                de.setIdDetailsInvoiceSell(rs.getInt("idDetailsInvoiceSell"));
//                de.setIdProduct(rs.getInt("idProduct"));
//                de.setIdInvoiceSell(rs.getInt("idInvoiceSell"));
                de.setQuantity(rs.getInt("quatity"));
                de.setPrice(rs.getInt("price"));
                de.setValueColor(rs.getString("valueColor"));
                de.setValueMaterial(rs.getString("valueMaterial"));
                de.setValueSize(rs.getString("valueSize"));
                de.setNameProduct(rs.getString("nameProduct"));
                de.setNameCustomer(rs.getString("name"));
//                de.setValueVoucher(rs.getInt("value"));
                list.add(de);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    public List<DetailInvoiceSell> selectByIdInvoice(int id) {
        String sql = "SELECT idDetailsInvoiceSELL, nameProduct, name, valueSize, valueColor, valueMaterial, detailsInvoiceSELL.quatity, detailsInvoiceSELL.price  FROM dbo.detailsInvoiceSELL \n"
                + "JOIN dbo.InvoiceSell ON InvoiceSell.idInvoiceSell = detailsInvoiceSELL.idInvoiceSell\n"
                + "JOIN dbo.Customer ON Customer.idCustomer = InvoiceSell.idCustomer\n"
                + "JOIN dbo.detailsProduct ON detailsProduct.idPrDeltails = detailsInvoiceSELL.idPrDetails\n"
                + "JOIN dbo.Products ON Products.idProduct = detailsProduct.idProduct JOIN dbo.Size ON Size.idSize = detailsProduct.idSize\n"
                + "JOIN dbo.Color ON Color.idColor = detailsProduct.idColor JOIN dbo.Material ON Material.idMaterial = detailsProduct.idMaterial\n"
                + "WHERE detailsInvoiceSELL.idInvoiceSell = ?";
        return selectBySql(sql, id);
    }
    

}
