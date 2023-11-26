import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.Cursor;
import java.awt.SystemColor;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.RowFilter;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.DefaultCellEditor;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableRowSorter;
import javax.swing.ImageIcon;
import javax.swing.border.LineBorder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OperativInform extends JFrame {

	private JPanel contentPane;
	private static JTable table;
	private JTextField textField;
	private JTextField textField_1;
	private JTextField textField_2;
	private JScrollPane scrollPane;
	private JTextField textField_3;
	private JTextField textField_4;
	private JTextField textField_5;
	private JTextField textField_6;
	private JTextField textField_7;
    
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					OperativInform frame = new OperativInform();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public OperativInform() {	
		
		// --- Окно --- //
		setForeground(SystemColor.control);
		setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
		setTitle("Оперативная информация об экспорте товара");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(0, 0, 1222, 498);
		setLocationRelativeTo(null);
		contentPane = new JPanel();
		contentPane.setForeground(Color.BLACK);
		contentPane.setBackground(Color.WHITE);
		setContentPane(contentPane);

		// --- Убрать подсветку кнопок --- //
		UIManager.put("Button.select", SystemColor.info);
		
		// --- CardLayout --- //
		CardLayout cl = new CardLayout();
		contentPane.setLayout(cl);
		
		///////////////////////////
		// --- Первая панель --- //
		///////////////////////////
		JPanel panelFirst = new JPanel();
		panelFirst.setForeground(Color.BLACK);
		panelFirst.setBackground(new Color(248, 248, 255));
		panelFirst.setLayout(null);
		contentPane.add(panelFirst, "1");
		
		// --- Таблица --- //
		scrollPane = new JScrollPane();
		scrollPane.setBorder(null);
		scrollPane.setForeground(Color.BLACK);
		scrollPane.setBackground(Color.WHITE);
		scrollPane.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		scrollPane.setBounds(25, 25, 1153, 270);
		
		table = new JTable(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"Товар", "Группа", "Объем", "Страна-экспортер", "Страна-импортер", "Дата отправки", "Дата прибытия"
			}
		));
		table.setForeground(Color.BLACK);
		table.setBackground(Color.WHITE);
		table.setSelectionBackground(new Color(216, 191, 216));
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		Font bigFont = new Font("Comic Sans MS", Font.PLAIN, 16);
		table.getTableHeader().setFont(bigFont);
		table.getTableHeader().setBackground(Color.WHITE);
		
		table.getColumnModel().getColumn(0).setPreferredWidth(190);
		table.getColumnModel().getColumn(1).setPreferredWidth(250);
		table.getColumnModel().getColumn(2).setPreferredWidth(100);
		table.getColumnModel().getColumn(3).setPreferredWidth(150);
		table.getColumnModel().getColumn(4).setPreferredWidth(200);
		table.getColumnModel().getColumn(5).setPreferredWidth(120);
		table.getColumnModel().getColumn(6).setPreferredWidth(125);

		table.setRowHeight(30);
		
		scrollPane.setViewportView(table);
		table.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		panelFirst.add(scrollPane);
		
		DefaultTableModel model = (DefaultTableModel)table.getModel();
		
		// --- Добавление в таблицу Combobox с данными из Excel --- //
		TovariExcel(true);
		GruppiExcel(true);
		ImporteriExcel(true);
		
		// --- Загрузка данных из Excel в таблицу--- //
		addWindowListener(new WindowAdapter() {
			@Override
			public void windowOpened(WindowEvent e) {
				FileInputStream file = null;
				XSSFWorkbook workbook = null;
				
				try {
					file = new FileInputStream(new File("C:\\Users\\79034\\Desktop\\Индивидуальное задание Java Волошиной Алины.xlsx"));
					workbook = new XSSFWorkbook (file);
				} catch (FileNotFoundException e1) {
					e1.printStackTrace();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
				XSSFSheet excelSheet = workbook.getSheetAt(0);
				
				try {
					for (int row = 1; row <= excelSheet.getLastRowNum(); row++) {
						XSSFRow excelRow = excelSheet.getRow(row);

		                XSSFCell excelTovar = excelRow.getCell(0);
		                XSSFCell excelGruppa = excelRow.getCell(1);
		                XSSFCell excelObem = excelRow.getCell(2);
		                XSSFCell excelStranaEksporter = excelRow.getCell(3);
		                XSSFCell excelStranaImporter = excelRow.getCell(4);
		                XSSFCell excelDataOtpravki = excelRow.getCell(5);
		                XSSFCell excelDataPribitiya = excelRow.getCell(6);
		                    
		                model.addRow(new Object[]{excelTovar, excelGruppa, excelObem, excelStranaEksporter, excelStranaImporter, excelDataOtpravki, 
		                		excelDataPribitiya});
		            }
					JOptionPane.showMessageDialog(null, "Загрузка прошла успешно");
		        } finally {
		        	try {
		        		if (workbook != null) {
		        			workbook.close();
		                }
		            } catch (IOException iOException) {
		            	JOptionPane.showMessageDialog(null, iOException.getMessage());
		            }
		        }
			}
		});
				
		// --- Кнопки --- //
		// --- Добавить --- //
		JButton btnNewButton = new JButton("");
		btnNewButton.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (5).png")));
		btnNewButton.setFocusPainted(false);
		btnNewButton.setFocusable(false);
		btnNewButton.setBackground(new Color(248, 248, 255));
		btnNewButton.setBorder(null);
		btnNewButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent arg0) {
				cl.show(contentPane, "2");
			}
		});
		btnNewButton.setBounds(446, 395, 104, 35);
		panelFirst.add(btnNewButton);
		
		// --- Удалить --- //
		JButton btnNewButton_1 = new JButton("");
		btnNewButton_1.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (6).png")));
		btnNewButton_1.setFocusPainted(false);
		btnNewButton_1.setFocusable(false);
		btnNewButton_1.setBackground(new Color(248, 248, 255));
		btnNewButton_1.setBorder(null);
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				int result = JOptionPane.showConfirmDialog(null,
						"Вы уверены, что хотите удалить эту запись?",
						getTitle(),
						JOptionPane.YES_NO_OPTION,
						JOptionPane.WARNING_MESSAGE);
				if (result == JOptionPane.YES_OPTION) {
					int getSelectedRowforDeletion = table.getSelectedRow();
					if (getSelectedRowforDeletion >= 0) {
						model.removeRow(getSelectedRowforDeletion);
						JOptionPane.showMessageDialog(null, "Запись удалена успешно");
					}
				}
			}
		});
		btnNewButton_1.setBounds(560, 395, 96, 35);
		panelFirst.add(btnNewButton_1);
		
		// --- Сохранить --- //
		JButton btnNewButton_2 = new JButton("");
		btnNewButton_2.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (7).png")));
		btnNewButton_2.setFocusable(false);
		btnNewButton_2.setFocusPainted(false);
		btnNewButton_2.setBorder(null);
		btnNewButton_2.setBackground(new Color(248, 248, 255));
		btnNewButton_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String path = "C:\\Users\\79034\\Desktop\\Индивидуальное задание Java Волошиной Алины.xlsx";
		        FileInputStream fileinp = null;
		        XSSFWorkbook workbook = null;
		        
				try {
					fileinp = new FileInputStream(path);
					workbook = new XSSFWorkbook(fileinp);
					} catch (FileNotFoundException e2) {
						e2.printStackTrace();
					} catch (IOException e2) {
						e2.printStackTrace();
					}
				
				int index = 0;
				XSSFSheet sheet = workbook.getSheet("Оперативная инф. об экс. товара");
				if(sheet != null)   {
				    index = workbook.getSheetIndex(sheet);
				    workbook.removeSheetAt(index);
				}
				
				XSSFSheet excelSheet = workbook.createSheet("Оперативная инф. об экс. товара");
				workbook.setSheetOrder("Оперативная инф. об экс. товара", 0);
				
				try {
					Row rowCol = excelSheet.createRow(0);
					for (int k = 0; k < table.getColumnCount(); k++) {
						Cell excel = rowCol.createCell(k);         
		                excel.setCellValue(table.getColumnName(k));
		                CellStyle style = workbook.createCellStyle();
		                style.setBorderTop(BorderStyle.THIN);
		                style.setBorderBottom(BorderStyle.THIN);
		                style.setBorderLeft(BorderStyle.THIN);
		                style.setBorderRight(BorderStyle.THIN);
		                XSSFFont font = workbook.createFont();
		                font.setFontName("Calibri");
		                font.setFontHeightInPoints((short) 12);
		                font.setBold(true);
		                style.setFont(font);
		                excel.setCellStyle(style);
		            }
					
					for (int i = 0; i < model.getRowCount(); i++) {
						XSSFRow excelRow = excelSheet.createRow(i + 1);
						for (int j = 0; j < model.getColumnCount(); j++) {
							XSSFCell excelCell = excelRow.createCell(j);         
							excelCell.setCellValue(model.getValueAt(i, j).toString());   
							excelSheet.autoSizeColumn(j);
						}
					}
				
					JOptionPane.showMessageDialog(null, "Сохранение прошло успешно");

				} finally {
					FileOutputStream fileOut = null;
					try {
						fileOut = new FileOutputStream(path);
						workbook.write(fileOut);
						fileOut.close();
						fileinp.close();
					} catch (FileNotFoundException e1) {
						e1.printStackTrace();
					} catch (IOException e1) {
						e1.printStackTrace();
					}
		        }
			}
		});
		btnNewButton_2.setBounds(666, 395, 112, 35);
		panelFirst.add(btnNewButton_2);
		
		// --- Поиск по названию товара --- //
		JLabel lblNewLabel_7 = new JLabel("");
		lblNewLabel_7.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (22).png")));
		lblNewLabel_7.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_7.setBorder(null);
		lblNewLabel_7.setBackground(new Color(248, 248, 255));
		lblNewLabel_7.setBounds(60, 306, 240, 35);
		panelFirst.add(lblNewLabel_7);
		
		textField_4 = new JTextField();
		textField_4.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				TableRowSorter<DefaultTableModel> tr = new TableRowSorter<DefaultTableModel>(model);
				table.setRowSorter(tr);
				tr.setRowFilter(RowFilter.regexFilter(textField_4.getText(), 0));
			}
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField_4.getText().length() >= 50)
					e.consume();
		    }
		});
		textField_4.setSelectionColor(new Color(230, 230, 250));
		textField_4.setSelectedTextColor(new Color(230, 230, 250));
		textField_4.setForeground(Color.BLACK);
		textField_4.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField_4.setDisabledTextColor(new Color(230, 230, 250));
		textField_4.setColumns(10);
		textField_4.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField_4.setBackground(Color.WHITE);
		textField_4.setBounds(310, 308, 220, 30);
		panelFirst.add(textField_4);
		
		// --- Поиск по названию страны-импортеру --- //
		JLabel lblNewLabel_8 = new JLabel("");
		lblNewLabel_8.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (23).png")));
		lblNewLabel_8.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_8.setBorder(null);
		lblNewLabel_8.setBackground(new Color(248, 248, 255));
		lblNewLabel_8.setBounds(591, 306, 334, 35);
		panelFirst.add(lblNewLabel_8);
		
		textField_5 = new JTextField();
		textField_5.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				TableRowSorter<DefaultTableModel> tr = new TableRowSorter<DefaultTableModel>(model);
				table.setRowSorter(tr);
				tr.setRowFilter(RowFilter.regexFilter(textField_5.getText(), 4));
			}
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField_5.getText().length() >= 35)
					e.consume();
		    }
		});
		textField_5.setSelectionColor(new Color(230, 230, 250));
		textField_5.setSelectedTextColor(new Color(230, 230, 250));
		textField_5.setForeground(Color.BLACK);
		textField_5.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField_5.setDisabledTextColor(new Color(230, 230, 250));
		textField_5.setColumns(10);
		textField_5.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField_5.setBackground(Color.WHITE);
		textField_5.setBounds(935, 308, 200, 30);
		panelFirst.add(textField_5);
		
		// --- Поиск по дате отправки --- //
		JLabel lblNewLabel_9 = new JLabel("");
		lblNewLabel_9.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (24).png")));
		lblNewLabel_9.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_9.setBorder(null);
		lblNewLabel_9.setBackground(new Color(248, 248, 255));
		lblNewLabel_9.setBounds(196, 352, 220, 35);
		panelFirst.add(lblNewLabel_9);
		
		textField_6 = new JTextField();
		textField_6.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				TableRowSorter<DefaultTableModel> tr = new TableRowSorter<DefaultTableModel>(model);
				table.setRowSorter(tr);
				tr.setRowFilter(RowFilter.regexFilter(textField_6.getText(), 5));
			}
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField_6.getText().length() >= 10)
					e.consume();
		    }
		});
		textField_6.setSelectionColor(new Color(230, 230, 250));
		textField_6.setSelectedTextColor(new Color(230, 230, 250));
		textField_6.setForeground(Color.BLACK);
		textField_6.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField_6.setDisabledTextColor(new Color(230, 230, 250));
		textField_6.setColumns(10);
		textField_6.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField_6.setBackground(Color.WHITE);
		textField_6.setBounds(426, 354, 150, 30);
		panelFirst.add(textField_6);
		
		// --- Поиск по дате прибытия --- //
		JLabel lblNewLabel_10 = new JLabel("");
		lblNewLabel_10.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (25).png")));
		lblNewLabel_10.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_10.setBorder(null);
		lblNewLabel_10.setBackground(new Color(248, 248, 255));
		lblNewLabel_10.setBounds(636, 352, 226, 35);
		panelFirst.add(lblNewLabel_10);
		
		textField_7 = new JTextField();
		textField_7.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				TableRowSorter<DefaultTableModel> tr = new TableRowSorter<DefaultTableModel>(model);
				table.setRowSorter(tr);
				tr.setRowFilter(RowFilter.regexFilter(textField_7.getText(), 6));
			}
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField_7.getText().length() >= 10)
					e.consume();
		    }
		});
		textField_7.setSelectionColor(new Color(230, 230, 250));
		textField_7.setSelectedTextColor(new Color(230, 230, 250));
		textField_7.setForeground(Color.BLACK);
		textField_7.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField_7.setDisabledTextColor(new Color(230, 230, 250));
		textField_7.setColumns(10);
		textField_7.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField_7.setBackground(Color.WHITE);
		textField_7.setBounds(872, 354, 150, 30);
		panelFirst.add(textField_7);
		
		///////////////////////////
		// --- Вторая панель --- //
		///////////////////////////
		JPanel panelSecond = new JPanel();
		panelSecond.setBackground(new Color(248, 248, 255));
		panelSecond.setLayout(null);
		contentPane.add(panelSecond, "2");

		// --- Label и TextField для добавления --- //
		// --- Товар --- //
		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (10).png")));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBorder(null);
		lblNewLabel.setBackground(new Color(248, 248, 255));
		lblNewLabel.setBounds(96, 141, 78, 35);
		panelSecond.add(lblNewLabel);
		
		JComboBox comboBox = new JComboBox();
		comboBox.setForeground(Color.BLACK);
		comboBox.setBorder(new LineBorder(new Color(147, 112, 219)));
		comboBox.setBackground(Color.WHITE);
		comboBox.setModel(new DefaultComboBoxModel(TovariExcel(false)));
		comboBox.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		comboBox.setBounds(184, 143, 220, 30);
		panelSecond.add(comboBox);
		
		// --- Группа --- //
		JLabel lblNewLabel_1 = new JLabel("");
		lblNewLabel_1.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (9).png")));
		lblNewLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_1.setBorder(null);
		lblNewLabel_1.setBackground(new Color(248, 248, 255));
		lblNewLabel_1.setBounds(464, 141, 88, 35);
		panelSecond.add(lblNewLabel_1);
		
		JComboBox comboBox_1 = new JComboBox();
		comboBox_1.setForeground(Color.BLACK);
		comboBox_1.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		comboBox_1.setBorder(new LineBorder(new Color(147, 112, 219)));
		comboBox_1.setBackground(Color.WHITE);
		comboBox_1.setModel(new DefaultComboBoxModel(GruppiExcel(false)));
		comboBox_1.setBounds(562, 143, 260, 30);
		panelSecond.add(comboBox_1);
		
		// --- Объем --- //
		JLabel lblNewLabel_2 = new JLabel("");
		lblNewLabel_2.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (16).png")));
		lblNewLabel_2.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_2.setBorder(null);
		lblNewLabel_2.setBackground(new Color(248, 248, 255));
		lblNewLabel_2.setBounds(882, 141, 84, 35);
		panelSecond.add(lblNewLabel_2);
		
		textField = new JTextField();
		textField.setSelectedTextColor(new Color(230, 230, 250));
		textField.setDisabledTextColor(new Color(230, 230, 250));
		textField.setSelectionColor(new Color(230, 230, 250));
		textField.setBackground(Color.WHITE);
		textField.addKeyListener((KeyListener) new KeyAdapter() {
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField.getText().length() >= 10)
					e.consume();
		    }
		});
		textField.setForeground(Color.BLACK);
		textField.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField.setColumns(10);
		textField.setBounds(976, 143, 130, 30);
		panelSecond.add(textField);
		
		// --- Страна-экспортер --- //
		JLabel lblNewLabel_3 = new JLabel("");
		lblNewLabel_3.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (17).png")));
		lblNewLabel_3.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_3.setBorder(null);
		lblNewLabel_3.setBackground(new Color(248, 248, 255));
		lblNewLabel_3.setBounds(204, 187, 168, 35);
		panelSecond.add(lblNewLabel_3);
		
		textField_1 = new JTextField();
		textField_1.setSelectionColor(new Color(230, 230, 250));
		textField_1.setSelectedTextColor(new Color(230, 230, 250));
		textField_1.setDisabledTextColor(new Color(230, 230, 250));
		textField_1.setBackground(Color.WHITE);
		textField_1.addKeyListener((KeyListener) new KeyAdapter() {
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField_1.getText().length() >= 50)
					e.consume();
		    }
		});
		textField_1.setForeground(Color.BLACK);
		textField_1.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField_1.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField_1.setColumns(10);
		textField_1.setBounds(382, 189, 200, 30);
		panelSecond.add(textField_1);
		
		// --- Страна-импортер --- //
		JLabel lblNewLabel_4 = new JLabel("");
		lblNewLabel_4.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (18).png")));
		lblNewLabel_4.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_4.setBorder(null);
		lblNewLabel_4.setBackground(new Color(248, 248, 255));
		lblNewLabel_4.setBounds(642, 187, 164, 35);
		panelSecond.add(lblNewLabel_4);
		
		JComboBox comboBox_2 = new JComboBox();
		comboBox_2.setForeground(Color.BLACK);
		comboBox_2.setModel(new DefaultComboBoxModel(ImporteriExcel(false)));
		comboBox_2.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		comboBox_2.setBorder(new LineBorder(new Color(147, 112, 219)));
		comboBox_2.setBackground(Color.WHITE);
		comboBox_2.setBounds(816, 189, 200, 30);
		panelSecond.add(comboBox_2);
		
		// --- Дата отправки --- //
		JLabel lblNewLabel_5 = new JLabel("");
		lblNewLabel_5.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (19).png")));
		lblNewLabel_5.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_5.setBorder(null);
		lblNewLabel_5.setBackground(new Color(248, 248, 255));
		lblNewLabel_5.setBounds(277, 232, 144, 35);
		panelSecond.add(lblNewLabel_5);
		
		textField_2 = new JTextField();
		textField_2.addKeyListener((KeyListener) new KeyAdapter() {
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField_2.getText().length() >= 10)
					e.consume();
		    }
		});
		textField_2.setSelectionColor(new Color(230, 230, 250));
		textField_2.setSelectedTextColor(new Color(230, 230, 250));
		textField_2.setForeground(Color.BLACK);
		textField_2.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField_2.setDisabledTextColor(new Color(230, 230, 250));
		textField_2.setColumns(10);
		textField_2.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField_2.setBackground(Color.WHITE);
		textField_2.setBounds(432, 234, 150, 30);
		panelSecond.add(textField_2);
		
		// --- Дата прибытия --- //
		JLabel lblNewLabel_6 = new JLabel("");
		lblNewLabel_6.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (20).png")));
		lblNewLabel_6.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_6.setBorder(null);
		lblNewLabel_6.setBackground(new Color(248, 248, 255));
		lblNewLabel_6.setBounds(642, 232, 148, 35);
		panelSecond.add(lblNewLabel_6);
		
		textField_3 = new JTextField();
		textField_3.addKeyListener((KeyListener) new KeyAdapter() {
			@Override
		    public void keyTyped(KeyEvent e) {
				if (textField_3.getText().length() >= 10)
					e.consume();
		    }
		});
		textField_3.setSelectionColor(new Color(230, 230, 250));
		textField_3.setSelectedTextColor(new Color(230, 230, 250));
		textField_3.setForeground(Color.BLACK);
		textField_3.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		textField_3.setDisabledTextColor(new Color(230, 230, 250));
		textField_3.setColumns(10);
		textField_3.setBorder(new LineBorder(new Color(147, 112, 219)));
		textField_3.setBackground(Color.WHITE);
		textField_3.setBounds(800, 234, 150, 30);
		panelSecond.add(textField_3);
		
		// --- Кнопка "ОК" --- //
		JButton btnNewButton_3 = new JButton("");
		btnNewButton_3.setIcon(new ImageIcon(OperativInform.class.getResource("/icon/button (21).png")));
		btnNewButton_3.setFocusPainted(false);
		btnNewButton_3.setFocusable(false);
		btnNewButton_3.setBorder(null);
		btnNewButton_3.setBackground(new Color(248, 248, 255));
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 model.addRow(new Object[]{
						 comboBox.getSelectedItem(),
						 comboBox_1.getSelectedItem(),
						 textField.getText(),
						 textField_1.getText(),
						 comboBox_2.getSelectedItem(),
						 textField_2.getText(),
						 textField_3.getText(),
                  });
				 JOptionPane.showMessageDialog(null, "Запись добавлена успешно");
				 cl.show(contentPane, "1");
			}
		});
		
		btnNewButton_3.setBounds(587, 278, 54, 35);
		panelSecond.add(btnNewButton_3);
		
		cl.show(contentPane, "1");
	}

	public static String[] TovariExcel(boolean chetka) {
		 FileInputStream file = null;
		 XSSFWorkbook workbook = null;
		
		 try {
			 file = new FileInputStream(new File("C:\\Users\\79034\\Desktop\\Индивидуальное задание Java Волошиной Алины.xlsx"));
			 workbook = new XSSFWorkbook (file);
		 } catch (FileNotFoundException e1) {
			 e1.printStackTrace();
		 } catch (IOException e) {
			 e.printStackTrace();
		 }
		 
		 XSSFSheet excelSheet = workbook.getSheetAt(1);
		 ArrayList<String> tovari = new ArrayList<String>();
			
		 try {
			 for(int i = 1; i <= excelSheet.getLastRowNum(); i++) {
				 Row row = excelSheet.getRow(i);
				 if(row != null) {
					 Cell cell = row.getCell(2);
					 if (cell != null) {
						 tovari.add(cell.getStringCellValue());
					 }
				 }
			 }
		
			 String[] array = (String[]) tovari.toArray(new String[0]);	
			
			 JComboBox comboBox = new JComboBox();
			 comboBox.setEditable(true);
			 comboBox.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
			 comboBox.setModel(new DefaultComboBoxModel(array));
		 
			 if (chetka) {
				 table.getColumnModel().getColumn(0).setCellEditor(new DefaultCellEditor(comboBox));
				 return null;
			 }
			 else
				 return array;
	 	
		 } finally {
			 try {
				 if (workbook != null) {
					 workbook.close();
				 }
			 } catch (IOException iOException) {
				 JOptionPane.showMessageDialog(null, iOException.getMessage());
			 }
		 }
	 }
	 public static String[] GruppiExcel(boolean chetka) {
		 FileInputStream file = null;
		 XSSFWorkbook workbook = null;
		
		 try {
			 file = new FileInputStream(new File("C:\\Users\\79034\\Desktop\\Индивидуальное задание Java Волошиной Алины.xlsx"));
			 workbook = new XSSFWorkbook (file);
		 } catch (FileNotFoundException e1) {
			 e1.printStackTrace();
		 } catch (IOException e) {
			 e.printStackTrace();
		 }
		 
		 XSSFSheet excelSheet = workbook.getSheetAt(2);
		 ArrayList<String> gruppi = new ArrayList<String>();
			
		 try {
			 for(int i = 1; i <= excelSheet.getLastRowNum(); i++) {
				 Row row = excelSheet.getRow(i);
				 if(row != null) {
					 Cell cell = row.getCell(0);
					 if (cell != null) {
						 gruppi.add(cell.getStringCellValue());
					 }
				 }
			 }
		
			 String[] array = (String[]) gruppi.toArray(new String[0]);	
			
			 JComboBox comboBox = new JComboBox();
			 comboBox.setEditable(true);
			 comboBox.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
			 comboBox.setModel(new DefaultComboBoxModel(array));
		 
			 if (chetka) {
				 table.getColumnModel().getColumn(1).setCellEditor(new DefaultCellEditor(comboBox));
				 return null;
			 }
			 else
				 return array;
	 	
		 } finally {
			 try {
				 if (workbook != null) {
					 workbook.close();
				 }
			 } catch (IOException iOException) {
				 JOptionPane.showMessageDialog(null, iOException.getMessage());
			 }
		 }
	 }
	 public static String[] ImporteriExcel(boolean chetka) {
		 FileInputStream file = null;
		 XSSFWorkbook workbook = null;
		
		 try {
			 file = new FileInputStream(new File("C:\\Users\\79034\\Desktop\\Индивидуальное задание Java Волошиной Алины.xlsx"));
			 workbook = new XSSFWorkbook (file);
		 } catch (FileNotFoundException e1) {
			 e1.printStackTrace();
		 } catch (IOException e) {
			 e.printStackTrace();
		 }
		 
		 XSSFSheet excelSheet = workbook.getSheetAt(3);
		 ArrayList<String> eksporteri = new ArrayList<String>();
			
		 try {
			 for(int i = 1; i <= excelSheet.getLastRowNum(); i++) {
				 Row row = excelSheet.getRow(i);
				 if(row != null) {
					 Cell cell = row.getCell(1);
					 if (cell != null) {
						 eksporteri.add(cell.getStringCellValue());
					 }
				 }
			 }
		
			 String[] array = (String[]) eksporteri.toArray(new String[0]);	
			
			 JComboBox comboBox = new JComboBox();
			 comboBox.setEditable(true);
			 comboBox.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
			 comboBox.setModel(new DefaultComboBoxModel(array));
		 
			 if (chetka) {
				 table.getColumnModel().getColumn(4).setCellEditor(new DefaultCellEditor(comboBox));
				 return null;
			 }
			 else
				 return array;
	 	
		 } finally {
			 try {
				 if (workbook != null) {
					 workbook.close();
				 }
			 } catch (IOException iOException) {
				 JOptionPane.showMessageDialog(null, iOException.getMessage());
			 }
		 }
	 }
}
