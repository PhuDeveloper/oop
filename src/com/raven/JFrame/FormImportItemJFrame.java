
package com.raven.JFrame;

import DAO.ColorDAO;
import DAO.MaterialDAO;
import DAO.ProductItemDAO;
import DAO.ProductItemImageDAO;
import DAO.ProductsDAO;
import DAO.SizeDAO;
import com.tpd.Validate.Validate;
import com.tpd.entity.Color;
import com.tpd.entity.Material;
import com.tpd.entity.ProductItem;
import com.tpd.entity.ProductItemImage;
import com.tpd.entity.Products;
import com.tpd.entity.Size;
import com.tpd.utils.MsgBox;
import com.tpd.utils.XImage;
import com.raven.form.FormItems;
import com.raven.form.FormListEmpolyee;
import com.raven.form.MainForm;
import java.awt.Component;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import javax.swing.DefaultComboBoxModel;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.table.DefaultTableModel;


public class FormImportItemJFrame extends javax.swing.JFrame {

    
    public FormImportItemJFrame() {
        initComponents();

        setLocationRelativeTo(null);
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        
        statusForm();
        fillComboboxSize();
        fillComboboxColor();
        fillComboboxMaterial();
        fillComboboxProduct();
        btnDelete.setEnabled(false);
    }

    public void statusForm() {
        btnAddMaterial.setVisible(false);
        btnColorAdd.setVisible(false);
        btnEditColor.setVisible(false);
        btnEditMaterial.setVisible(false);
        btnEditSize.setVisible(false);
        btnAddSize.setVisible(false);
        txtColorAdd.setVisible(false);
        txtMaterialAdd.setVisible(false);
        txtSizeAdd.setVisible(false);
        lblColor.setVisible(false);
        lblMaterialAdd.setVisible(false);
        lblSizeAdd.setVisible(false);
        lblMaterialAdd.setVisible(false);
        lblPrice.setVisible(false);
    }
    
    SizeDAO sDao = new SizeDAO();
    ColorDAO cDao = new ColorDAO();
    MaterialDAO mDao = new MaterialDAO();
    ProductsDAO productDAO = new ProductsDAO();
    ProductItemDAO productItemDAO = new ProductItemDAO();
    ProductItemImageDAO productItemImageDAO = new ProductItemImageDAO();

    List<Size> listSize = sDao.selectAll();
    List<Material> listMaterial = mDao.selectAll();
    List<Color> listColor = cDao.selectAll();

    public void fillComboboxSize() {
        List<Size> listSize = sDao.selectAll();
        DefaultComboBoxModel model = (DefaultComboBoxModel) cbbSize.getModel();
        cbbSize.removeAllItems();
        for (Size s : listSize) {
            model.addElement(s);
        }
    }

    public void fillComboboxColor() {

        DefaultComboBoxModel model = (DefaultComboBoxModel) cbbColor.getModel();
        cbbColor.removeAllItems();
        List<Color> listColor = cDao.selectAll();
        for (Color c : listColor) {
            model.addElement(c);
        }
    }

    public void fillComboboxMaterial() {
        DefaultComboBoxModel model = (DefaultComboBoxModel) cbbMaterial.getModel();
        cbbMaterial.removeAllItems();
        List<Material> listMaterial = mDao.selectAll();

        for (Material c : listMaterial) {
            model.addElement(c);
        }
    }

    public void showSize() {
        Size s = (Size) cbbSize.getSelectedItem();
        if (s == null) {
            return;
        } else {
            txtSizeAdd.setText(s.getValueSize());
        }
    }

    public void showColor() {
        Color c = (Color) cbbColor.getSelectedItem();
        if (c == null) {
            return;
        } else {
            txtColorAdd.setText(c.getValueColor());
        }
    }

    public void showMaterial() {
        Material m = (Material) cbbMaterial.getSelectedItem();
        if (m == null) {
            return;
        } else {
            txtMaterialAdd.setText(m.getValueMaterial());
        }
    }

    Size getFormSize() {
        Size s = new Size();
        s.setValueSize(txtSizeAdd.getText());
        return s;
    }

    Color getFormColor() {
        Color c = new Color();
        c.setValueColor(txtColorAdd.getText());
        return c;
    }

    Material getFormMaterial() {
        Material m = new Material();
        m.setValueMaterial(txtMaterialAdd.getText());
        return m;
    }

    public void insertSize() {
        Size s = getFormSize();
        List<Size> list = sDao.selectAll();
        if (!Validate.checkEmpty(lblSizeAdd, txtSizeAdd, "Kh??ng b??? tr???ng size")) {
            lblSizeAdd.setVisible(true);
            return;
        }
        for (Size size : list) {
            if (txtSizeAdd.getText().equalsIgnoreCase(size.getValueSize())) {
                lblSizeAdd.setVisible(true);
                lblSizeAdd.setText("Size ???? c?? !!!");
                txtSizeAdd.setText("");
                return;
            }
        }
        sDao.insert(s);
        lblSizeAdd.setVisible(false);
        btnAddSize.setVisible(false);
        btnEditSize.setVisible(false);
        txtSizeAdd.setVisible(false);
        MsgBox.alert(this, "Th??m Th??nh c??ng");
        fillComboboxSize();
    }

    public void updateSize() {
        Size c = (Size) cbbSize.getSelectedItem();
        c.setValueSize(txtSizeAdd.getText());
        try {
            if (!Validate.checkEmpty(lblSizeAdd, txtSizeAdd, "Ch??a nh???p Size!")) {
                return;
            } else {
                sDao.update(c);
                MsgBox.alert(this, "S???a ?????i th??nh c??ng!!");
                fillComboboxColor();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void insertColor() {
        Color c = getFormColor();
        List<Color> list = cDao.selectAll();
        if (!Validate.checkEmpty(lblColor, txtColorAdd, "Kh??ng b??? tr???ng color")) {
            lblColor.setVisible(true);
            return;
        }
        for (Color color : list) {
            if (txtColorAdd.getText().equalsIgnoreCase(color.getValueColor())) {
                lblColor.setVisible(true);
                lblColor.setText("Color ???? c?? !!!");
                txtColorAdd.setText("");
                return;
            }
        }
        cDao.insert(c);
        lblColor.setVisible(false);
        btnColorAdd.setVisible(false);
        btnEditColor.setVisible(false);
        txtColorAdd.setVisible(false);
        MsgBox.alert(this, "Th??m Th??nh c??ng");
        fillComboboxColor();
    }

    public void updateColor() {
        Color c = (Color) cbbColor.getSelectedItem();
        c.setValueColor(txtColorAdd.getText());
        try {
            if (!Validate.checkEmpty(lblColor, txtColorAdd, "Ch??a nh???p m??u!")) {
                return;
            } else {
                cDao.update(c);
                MsgBox.alert(this, "S???a ?????i th??nh c??ng!!");
                fillComboboxColor();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void insertMaterial() {
        Material c = getFormMaterial();
        List<Material> list = mDao.selectAll();
        if (!Validate.checkEmpty(lblMaterialAdd, txtMaterialAdd, "Kh??ng b??? tr???ng ch???t li???u")) {
            lblMaterialAdd.setVisible(true);
            return;
        }
        for (Material color : list) {
            if (txtMaterialAdd.getText().equalsIgnoreCase(color.getValueMaterial())) {
                lblMaterialAdd.setVisible(true);
                lblMaterialAdd.setText("Ch???t li???u ???? c?? !!!");
                txtMaterialAdd.setText("");
                return;
            }
        }
        mDao.insert(c);
        lblMaterialAdd.setVisible(false);
        btnAddMaterial.setVisible(false);
        btnEditMaterial.setVisible(false);
        txtMaterialAdd.setVisible(false);
        MsgBox.alert(this, "Th??m Th??nh c??ng");
        fillComboboxMaterial();
    }

    public void updateMaterial() {
        Material c = (Material) cbbMaterial.getSelectedItem();
        c.setValueMaterial(txtMaterialAdd.getText());
        try {
            if (!Validate.checkEmpty(lblMaterialAdd, txtMaterialAdd, "Ch??a nh???p ch???t li???u!")) {
                return;
            } else {
                mDao.update(c);
                MsgBox.alert(this, "S???a ?????i th??nh c??ng!!");
                fillComboboxMaterial();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    List<Products> listProducts = productDAO.selectAll();

    public void fillComboboxProduct() {
        DefaultComboBoxModel model = (DefaultComboBoxModel) cbbProduct.getModel();
        cbbProduct.removeAllItems();

        for (Products p : listProducts) {
            model.addElement(p);
        }
    }

    public ProductItem getFormProductItem() {
        Material m = (Material) cbbMaterial.getSelectedItem();
        Size s = (Size) cbbSize.getSelectedItem();
        Color c = (Color) cbbColor.getSelectedItem();
        Products p = (Products) cbbProduct.getSelectedItem();
        ProductItem pro = new ProductItem();
        pro.setPrice(Float.parseFloat(txtPrice.getText()));
        pro.setIdSize(s.getIdSize());
        pro.setIdColor(c.getIdColor());
        pro.setIdMaterial(m.getIdMaterial());
        pro.setIdProduct(p.getIdProduct());
        pro.setQuantity(0);
        pro.setStatus(true);
        pro.setSize(s.getValueSize());
        pro.setColor(c.getValueColor());
        pro.setMaterial(m.getValueMaterial());
        pro.setProductName(p.getNameProduct());
        return pro;
    }

    List<ProductItem> list = new ArrayList<>();

    public boolean checkItem() {
        List<ProductItem> listItem = productItemDAO.selectAll();
        ProductItem newProductItem = getFormProductItem();

        for (ProductItem item : listItem) {
            if (item.getIdProduct() == newProductItem.getIdProduct()//C??ng 1 s???n ph???m
                    && item.getIdColor() == newProductItem.getIdColor()//C??ng 1 m??u
                    && item.getIdMaterial() == newProductItem.getIdMaterial()//C??ng 1 ch???t li???u  
                    && item.getIdSize() == newProductItem.getIdSize())//C??ng 1 size
            {
                MsgBox.alert(this, "M???t h??ng ???? c?? vui l??ng th??m m???t h??ng kh??c");
                return false;
            }
        }
        for (ProductItem item : list) {
            if (item.getIdProduct() == newProductItem.getIdProduct()//C??ng 1 s???n ph???m
                    && item.getIdColor() == newProductItem.getIdColor()//C??ng 1 m??u
                    && item.getIdMaterial() == newProductItem.getIdMaterial()//C??ng 1 ch???t li???u  
                    && item.getIdSize() == newProductItem.getIdSize())//C??ng 1 size
            {
                MsgBox.alert(this, "M???t h??ng v???a ???????c l??u t???m vui l??ng th??m m???t h??ng kh??c");
                return false;
            }
        }
        return true;
    }

    public void fillTableTemp() {
        if (!Validate.checkEmpty(lblPrice, txtPrice, "Kh??ng b??? tr???ng gi?? ti???n")) {
            lblPrice.setVisible(true);
            return;
        } else if (checkItem() == false) {
            return;
        } else {
            DefaultTableModel model = (DefaultTableModel) tableColumn1.getModel();
            ProductItem p = getFormProductItem();
            Object[] obj = new Object[]{p.getProductName(), p.getPrice(), p.getSize(), p.getColor(), p.getMaterial()};
            model.addRow(obj);
            list.add(p);

        }
    }

    public void insertProductItem() {
        int count = tableColumn1.getRowCount();
        if (count > 0) {
            for (int i = 0; i < list.size(); i++) {
                productItemDAO.insert(list.get(i));

                txtPrice.setText("");
            }
//            fillTableProducts();
            MsgBox.alert(this, "Th??m " + count + " m???t h??ng th??nh c??ng");
            list.clear();
            DefaultTableModel model = (DefaultTableModel) tableColumn1.getModel();
            model.setRowCount(0);
            txtPrice.setText("");
            new MainForm().showForm(new FormItems());
            this.dispose();
        } else {
            MsgBox.alert(this, "B???n ch??a th??m m???t h??ng n??o c???");
        }

    }

  

    public void deleteRowInTableTemp() {
        int row = tableColumn1.getSelectedRow();
        DefaultTableModel model = (DefaultTableModel) tableColumn1.getModel();

        for (int i = 0; i < tableColumn1.getRowCount(); i++) {
            if (row == i) {
                model.removeRow(row);
                list.remove(list.get(i));
                MsgBox.alert(this, "X??a m???t h??ng th??nh c??ng !");
                btnDelete.setEnabled(false);
                return;
            }
        }
    }

   




    public void addEvenFillTable(ActionListener evt) {
        btnAddProductItem.addActionListener(evt);
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel4 = new javax.swing.JPanel();
        jPanel5 = new javax.swing.JPanel();
        btnAddTemp = new com.raven.suportSwing.MyButton();
        btnDelete = new com.raven.suportSwing.MyButton();
        btnAddProductItem = new com.raven.suportSwing.MyButton();
        myButton7 = new com.raven.suportSwing.MyButton();
        jPanel1 = new javax.swing.JPanel();
        jLabel2 = new javax.swing.JLabel();
        jPanel3 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tableColumn1 = new com.raven.suportSwing.TableColumn();
        scrollBar1 = new com.raven.suportSwing.ScrollBar();
        jPanel2 = new javax.swing.JPanel();
        cbbProduct = new com.raven.suportSwing.Combobox();
        txtPrice = new com.raven.suportSwing.TextField();
        jPanel6 = new javax.swing.JPanel();
        btn = new com.raven.suportSwing.MyButton();
        txtColorAdd = new com.raven.suportSwing.TextField();
        btnEditColor = new com.raven.suportSwing.MyButton();
        cbbColor = new com.raven.suportSwing.Combobox();
        btnAddSize = new com.raven.suportSwing.MyButton();
        txtMaterialAdd = new com.raven.suportSwing.TextField();
        cbbSize = new com.raven.suportSwing.Combobox();
        cbbMaterial = new com.raven.suportSwing.Combobox();
        btnEditSize = new com.raven.suportSwing.MyButton();
        myButton11 = new com.raven.suportSwing.MyButton();
        txtSizeAdd = new com.raven.suportSwing.TextField();
        btnEditMaterial = new com.raven.suportSwing.MyButton();
        myButton10 = new com.raven.suportSwing.MyButton();
        btnColorAdd = new com.raven.suportSwing.MyButton();
        btnAddMaterial = new com.raven.suportSwing.MyButton();
        lblSizeAdd = new javax.swing.JLabel();
        lblColor = new javax.swing.JLabel();
        lblMaterialAdd = new javax.swing.JLabel();
        lblPrice = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setAlwaysOnTop(true);
        setBackground(new java.awt.Color(255, 255, 255));
        setUndecorated(true);

        jPanel4.setBackground(new java.awt.Color(255, 255, 255));

        jPanel5.setBackground(java.awt.Color.white);

        btnAddTemp.setText("L??u t???m");
        btnAddTemp.setRadius(20);
        btnAddTemp.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAddTempActionPerformed(evt);
            }
        });

        btnDelete.setText("Xo??");
        btnDelete.setRadius(20);
        btnDelete.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnDeleteActionPerformed(evt);
            }
        });

        btnAddProductItem.setText("Ho??n Th??nh");
        btnAddProductItem.setRadius(20);
        btnAddProductItem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAddProductItemActionPerformed(evt);
            }
        });

        myButton7.setText("Hu???");
        myButton7.setRadius(20);
        myButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                myButton7ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(btnAddProductItem, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(myButton7, javax.swing.GroupLayout.PREFERRED_SIZE, 66, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(btnDelete, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btnAddTemp, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addGap(59, 59, 59)
                .addComponent(btnAddTemp, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 462, Short.MAX_VALUE)
                .addComponent(btnDelete, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(btnAddProductItem, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(208, 208, 208)
                .addComponent(myButton7, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        jPanel1.setBackground(java.awt.Color.white);

        jLabel2.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel2.setText("Th??m m???t h??ng");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 126, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(153, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, 37, Short.MAX_VALUE))
        );

        jPanel3.setBackground(java.awt.Color.white);

        tableColumn1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "T??n s???n ph???m", "Gi?? B??n", "Size", "M??u S???c", "Ch???t li???u"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        tableColumn1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tableColumn1MouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(tableColumn1);
        if (tableColumn1.getColumnModel().getColumnCount() > 0) {
            tableColumn1.getColumnModel().getColumn(0).setResizable(false);
            tableColumn1.getColumnModel().getColumn(1).setResizable(false);
            tableColumn1.getColumnModel().getColumn(2).setResizable(false);
            tableColumn1.getColumnModel().getColumn(3).setResizable(false);
            tableColumn1.getColumnModel().getColumn(4).setResizable(false);
        }

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 927, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(scrollBar1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 292, Short.MAX_VALUE)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGap(41, 41, 41)
                .addComponent(scrollBar1, javax.swing.GroupLayout.PREFERRED_SIZE, 149, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(102, Short.MAX_VALUE))
        );

        jPanel2.setBackground(java.awt.Color.white);
        jPanel2.setBorder(javax.swing.BorderFactory.createTitledBorder(null, "Chi ti???t s???n ph???m", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Tahoma", 1, 11))); // NOI18N

        cbbProduct.setLabeText("T??n s???n ph???m");
        cbbProduct.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cbbProductActionPerformed(evt);
            }
        });

        txtPrice.setLabelText("Gi?? B??n");
        txtPrice.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                txtPriceFocusGained(evt);
            }
        });

        jPanel6.setBackground(new java.awt.Color(255, 255, 255));

        btn.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/raven/icon/Create.png"))); // NOI18N
        btn.setRadius(20);
        btn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnActionPerformed(evt);
            }
        });

        txtColorAdd.setLabelText("M??u Th??m");
        txtColorAdd.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                txtColorAddFocusGained(evt);
            }
        });

        btnEditColor.setText("S???a");
        btnEditColor.setMaximumSize(new java.awt.Dimension(59, 23));
        btnEditColor.setMinimumSize(new java.awt.Dimension(59, 23));
        btnEditColor.setPreferredSize(new java.awt.Dimension(59, 23));
        btnEditColor.setRadius(20);
        btnEditColor.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEditColorActionPerformed(evt);
            }
        });

        cbbColor.setLabeText("M??u s???c");
        cbbColor.setPreferredSize(new java.awt.Dimension(30, 40));
        cbbColor.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cbbColorActionPerformed(evt);
            }
        });

        btnAddSize.setText("Th??m");
        btnAddSize.setRadius(20);
        btnAddSize.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAddSizeActionPerformed(evt);
            }
        });

        txtMaterialAdd.setLabelText("Ch???t li???u Th??m");
        txtMaterialAdd.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                txtMaterialAddFocusGained(evt);
            }
        });
        txtMaterialAdd.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtMaterialAddActionPerformed(evt);
            }
        });

        cbbSize.setLabeText("Size");
        cbbSize.setPreferredSize(new java.awt.Dimension(30, 40));
        cbbSize.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                cbbSizeMouseClicked(evt);
            }
        });
        cbbSize.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cbbSizeActionPerformed(evt);
            }
        });

        cbbMaterial.setLabeText("Ch???t li???u");
        cbbMaterial.setPreferredSize(new java.awt.Dimension(30, 40));
        cbbMaterial.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cbbMaterialActionPerformed(evt);
            }
        });

        btnEditSize.setText("S???a");
        btnEditSize.setMaximumSize(new java.awt.Dimension(59, 23));
        btnEditSize.setMinimumSize(new java.awt.Dimension(59, 23));
        btnEditSize.setPreferredSize(new java.awt.Dimension(59, 23));
        btnEditSize.setRadius(20);
        btnEditSize.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEditSizeActionPerformed(evt);
            }
        });

        myButton11.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/raven/icon/Create.png"))); // NOI18N
        myButton11.setRadius(20);
        myButton11.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                myButton11ActionPerformed(evt);
            }
        });

        txtSizeAdd.setLabelText("Size Th??m");
        txtSizeAdd.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                txtSizeAddFocusGained(evt);
            }
        });
        txtSizeAdd.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtSizeAddActionPerformed(evt);
            }
        });

        btnEditMaterial.setText("S???a");
        btnEditMaterial.setMaximumSize(new java.awt.Dimension(59, 23));
        btnEditMaterial.setMinimumSize(new java.awt.Dimension(59, 23));
        btnEditMaterial.setPreferredSize(new java.awt.Dimension(59, 23));
        btnEditMaterial.setRadius(20);
        btnEditMaterial.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEditMaterialActionPerformed(evt);
            }
        });

        myButton10.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/raven/icon/Create.png"))); // NOI18N
        myButton10.setRadius(20);
        myButton10.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                myButton10ActionPerformed(evt);
            }
        });

        btnColorAdd.setText("Th??m");
        btnColorAdd.setRadius(20);
        btnColorAdd.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnColorAddActionPerformed(evt);
            }
        });

        btnAddMaterial.setText("Th??m");
        btnAddMaterial.setRadius(20);
        btnAddMaterial.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAddMaterialActionPerformed(evt);
            }
        });

        lblSizeAdd.setFont(new java.awt.Font("Tahoma", 2, 11)); // NOI18N
        lblSizeAdd.setForeground(new java.awt.Color(255, 0, 0));

        lblColor.setFont(new java.awt.Font("Tahoma", 2, 11)); // NOI18N
        lblColor.setForeground(new java.awt.Color(255, 0, 0));

        lblMaterialAdd.setFont(new java.awt.Font("Tahoma", 2, 11)); // NOI18N
        lblMaterialAdd.setForeground(new java.awt.Color(255, 51, 0));
        lblMaterialAdd.setText("jLabel3");

        javax.swing.GroupLayout jPanel6Layout = new javax.swing.GroupLayout(jPanel6);
        jPanel6.setLayout(jPanel6Layout);
        jPanel6Layout.setHorizontalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel6Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel6Layout.createSequentialGroup()
                        .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(lblColor, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel6Layout.createSequentialGroup()
                                .addComponent(cbbColor, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(myButton10, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel6Layout.createSequentialGroup()
                                .addComponent(cbbSize, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(btn, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel6Layout.createSequentialGroup()
                                .addComponent(btnAddMaterial, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(btnEditMaterial, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel6Layout.createSequentialGroup()
                                .addComponent(btnAddSize, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(btnEditSize, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                .addComponent(lblSizeAdd, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(txtSizeAdd, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 192, Short.MAX_VALUE))
                            .addGroup(jPanel6Layout.createSequentialGroup()
                                .addComponent(btnColorAdd, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(btnEditColor, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel6Layout.createSequentialGroup()
                                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(lblMaterialAdd, javax.swing.GroupLayout.DEFAULT_SIZE, 201, Short.MAX_VALUE)
                                    .addGroup(jPanel6Layout.createSequentialGroup()
                                        .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtMaterialAdd, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(cbbMaterial, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(0, 9, Short.MAX_VALUE)))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(myButton11, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addContainerGap(15, Short.MAX_VALUE))
                    .addGroup(jPanel6Layout.createSequentialGroup()
                        .addComponent(txtColorAdd, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
        );
        jPanel6Layout.setVerticalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel6Layout.createSequentialGroup()
                .addGap(8, 8, 8)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btn, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(cbbSize, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtSizeAdd, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(lblSizeAdd, javax.swing.GroupLayout.PREFERRED_SIZE, 12, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnAddSize, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnEditSize, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(myButton10, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(cbbColor, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtColorAdd, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(lblColor, javax.swing.GroupLayout.PREFERRED_SIZE, 12, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btnColorAdd, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnEditColor, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel6Layout.createSequentialGroup()
                        .addGap(22, 22, 22)
                        .addComponent(myButton11, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel6Layout.createSequentialGroup()
                        .addGap(18, 18, 18)
                        .addComponent(cbbMaterial, javax.swing.GroupLayout.PREFERRED_SIZE, 36, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtMaterialAdd, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(lblMaterialAdd)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnEditMaterial, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnAddMaterial, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap())
        );

        lblPrice.setForeground(new java.awt.Color(225, 0, 0));
        lblPrice.setText("jLabel3");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(124, 124, 124)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(cbbProduct, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtPrice, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(lblPrice, javax.swing.GroupLayout.PREFERRED_SIZE, 192, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(265, 265, 265)
                .addComponent(jPanel6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(87, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(31, 31, 31)
                .addComponent(cbbProduct, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtPrice, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(lblPrice)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 1078, Short.MAX_VALUE)
            .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanel4Layout.createSequentialGroup()
                    .addGap(5, 5, 5)
                    .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                    .addComponent(jPanel5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(20, Short.MAX_VALUE)))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 851, Short.MAX_VALUE)
            .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanel4Layout.createSequentialGroup()
                    .addContainerGap()
                    .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel4Layout.createSequentialGroup()
                            .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                            .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addComponent(jPanel5, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addContainerGap()))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel4, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnActionPerformed
        // TODO add your handling code here:
        if (txtSizeAdd.isVisible()) {
            txtSizeAdd.setVisible(false);
            btnAddSize.setVisible(false);
            btnEditSize.setVisible(false);
            lblSizeAdd.setText("");
        } else {
            txtSizeAdd.setVisible(true);
            btnAddSize.setVisible(true);
            btnEditSize.setVisible(true);
            showSize();
        }

    }//GEN-LAST:event_btnActionPerformed

    private void myButton10ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_myButton10ActionPerformed
        // TODO add your handling code here:
        if (txtColorAdd.isVisible()) {
            btnColorAdd.setVisible(false);
            txtColorAdd.setVisible(false);
            btnEditColor.setVisible(false);
            lblColor.setText("");
        } else {
            btnColorAdd.setVisible(true);
            txtColorAdd.setVisible(true);
            btnEditColor.setVisible(true);
            showColor();
        }
    }//GEN-LAST:event_myButton10ActionPerformed

    private void myButton11ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_myButton11ActionPerformed
        // TODO add your handling code here:
        if (txtMaterialAdd.isVisible()) {
            txtMaterialAdd.setVisible(false);
            btnAddMaterial.setVisible(false);
            btnEditMaterial.setVisible(false);
            lblMaterialAdd.setText("");
        } else {
            txtMaterialAdd.setVisible(true);
            btnAddMaterial.setVisible(true);
            btnEditMaterial.setVisible(true);
            showMaterial();
        }
    }//GEN-LAST:event_myButton11ActionPerformed

    private void cbbColorActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cbbColorActionPerformed
        // TODO add your handling code here:
        showColor();
    }//GEN-LAST:event_cbbColorActionPerformed

    private void txtMaterialAddActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtMaterialAddActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtMaterialAddActionPerformed

    private void btnEditMaterialActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEditMaterialActionPerformed
        // TODO add your handling code here:
        updateMaterial();
    }//GEN-LAST:event_btnEditMaterialActionPerformed

    private void cbbSizeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cbbSizeActionPerformed
        // TODO add your handling code here:
        showSize();
    }//GEN-LAST:event_cbbSizeActionPerformed

    private void btnAddSizeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAddSizeActionPerformed
        // TODO add your handling code here:
        insertSize();
    }//GEN-LAST:event_btnAddSizeActionPerformed

    private void cbbSizeMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_cbbSizeMouseClicked
        // TODO add your handling code here:
    }//GEN-LAST:event_cbbSizeMouseClicked

    private void btnColorAddActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnColorAddActionPerformed
        // TODO add your handling code here:
        insertColor();
    }//GEN-LAST:event_btnColorAddActionPerformed

    private void btnEditColorActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEditColorActionPerformed
        // TODO add your handling code here:
        updateColor();
    }//GEN-LAST:event_btnEditColorActionPerformed

    private void btnAddMaterialActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAddMaterialActionPerformed
        // TODO add your handling code here:
        insertMaterial();
    }//GEN-LAST:event_btnAddMaterialActionPerformed

    private void cbbMaterialActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cbbMaterialActionPerformed
        // TODO add your handling code here:
        showMaterial();
    }//GEN-LAST:event_cbbMaterialActionPerformed

    private void txtPriceFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtPriceFocusGained
        lblPrice.setVisible(false);
    }//GEN-LAST:event_txtPriceFocusGained

    private void btnAddTempActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAddTempActionPerformed
        fillTableTemp();
    }//GEN-LAST:event_btnAddTempActionPerformed


    private void btnAddProductItemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAddProductItemActionPerformed

    }//GEN-LAST:event_btnAddProductItemActionPerformed


    private void tableColumn1MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tableColumn1MouseClicked

        btnDelete.setEnabled(true);
    }//GEN-LAST:event_tableColumn1MouseClicked

    private void cbbProductActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cbbProductActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_cbbProductActionPerformed

    private void btnEditSizeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEditSizeActionPerformed
        // TODO add your handling code here:
        updateSize();
    }//GEN-LAST:event_btnEditSizeActionPerformed

    private void btnDeleteActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnDeleteActionPerformed
        deleteRowInTableTemp();
    }//GEN-LAST:event_btnDeleteActionPerformed

    private void myButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_myButton7ActionPerformed
        // TODO add your handling code here:
        this.dispose();
        txtPrice.setText("");
        cbbColor.setSelectedIndex(0);
        cbbMaterial.setSelectedIndex(0);
        cbbProduct.setSelectedIndex(0);
        cbbSize.setSelectedIndex(0);
        DefaultTableModel model = (DefaultTableModel) tableColumn1.getModel();
        model.setRowCount(0);
    }//GEN-LAST:event_myButton7ActionPerformed

    private void txtSizeAddFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtSizeAddFocusGained
        // TODO add your handling code here:
//        lblSizeAdd.setText("");
    }//GEN-LAST:event_txtSizeAddFocusGained

    private void txtColorAddFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtColorAddFocusGained
        // TODO add your handling code here:
//        lblColor.setText("");
    }//GEN-LAST:event_txtColorAddFocusGained

    private void txtMaterialAddFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_txtMaterialAddFocusGained
        // TODO add your handling code here:
//        lblMaterialAdd.setText("");
    }//GEN-LAST:event_txtMaterialAddFocusGained

    private void txtSizeAddActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtSizeAddActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtSizeAddActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(FormImportItemJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(FormImportItemJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(FormImportItemJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(FormImportItemJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new FormImportItemJFrame().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private com.raven.suportSwing.MyButton btn;
    private com.raven.suportSwing.MyButton btnAddMaterial;
    private com.raven.suportSwing.MyButton btnAddProductItem;
    private com.raven.suportSwing.MyButton btnAddSize;
    private com.raven.suportSwing.MyButton btnAddTemp;
    private com.raven.suportSwing.MyButton btnColorAdd;
    private com.raven.suportSwing.MyButton btnDelete;
    private com.raven.suportSwing.MyButton btnEditColor;
    private com.raven.suportSwing.MyButton btnEditMaterial;
    private com.raven.suportSwing.MyButton btnEditSize;
    private com.raven.suportSwing.Combobox cbbColor;
    private com.raven.suportSwing.Combobox cbbMaterial;
    private com.raven.suportSwing.Combobox cbbProduct;
    private com.raven.suportSwing.Combobox cbbSize;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JPanel jPanel6;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JLabel lblColor;
    private javax.swing.JLabel lblMaterialAdd;
    private javax.swing.JLabel lblPrice;
    private javax.swing.JLabel lblSizeAdd;
    private com.raven.suportSwing.MyButton myButton10;
    private com.raven.suportSwing.MyButton myButton11;
    private com.raven.suportSwing.MyButton myButton7;
    private com.raven.suportSwing.ScrollBar scrollBar1;
    private com.raven.suportSwing.TableColumn tableColumn1;
    private com.raven.suportSwing.TextField txtColorAdd;
    private com.raven.suportSwing.TextField txtMaterialAdd;
    private com.raven.suportSwing.TextField txtPrice;
    private com.raven.suportSwing.TextField txtSizeAdd;
    // End of variables declaration//GEN-END:variables

}
