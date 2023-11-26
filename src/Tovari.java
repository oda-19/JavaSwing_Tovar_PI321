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
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.DefaultCellEditor;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.table.DefaultTableModel;
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

public class Tovari extends JFrame {

	private JPanel contentPane;
	private static JTable table;
	private JTextField textField;
	private JTextField textField_1;
	private JScrollPane scrollPane;
    
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Tovari frame = new Tovari();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public Tovari() {	
		
		// --- Окно --- //
		setForeground(SystemColor.control);
		setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
		setTitle("Список товаров");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(0, 0, 752, 406);
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
		scrollPane.setBounds(25, 25, 684, 270);
		
		table = new JTable(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"Номер товара", "Группа", "Товар"
			}
		));
		table.setForeground(Color.BLACK);
		table.setBackground(Color.WHITE);
		table.setSelectionBackground(new Color(216, 191, 216));
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		Font bigFont = new Font("Comic Sans MS", Font.PLAIN, 16);
		table.getTableHeader().setFont(bigFont);
		table.getTableHeader().setBackground(Color.WHITE);
		
		table.getColumnModel().getColumn(0).setPreferredWidth(115);
		table.getColumnModel().getColumn(1).setPreferredWidth(280);
		table.getColumnModel().getColumn(2).setPreferredWidth(270);

		table.setRowHeight(30);
		
		scrollPane.setViewportView(table);
		table.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		panelFirst.add(scrollPane);
		
		DefaultTableModel model = (DefaultTableModel)table.getModel();
		
		// --- Добавление в таблицу Combobox с данными из Excel --- //
		GruppiExcel(true);
		
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
				XSSFSheet excelSheet = workbook.getSheetAt(1);
				
				try {
					for (int row = 1; row <= excelSheet.getLastRowNum(); row++) {
						XSSFRow excelRow = excelSheet.getRow(row);

		                XSSFCell excelNomer = excelRow.getCell(0);
		                XSSFCell excelGruppa = excelRow.getCell(1);
		                XSSFCell excelTovar = excelRow.getCell(2);
		                    
		                model.addRow(new Object[]{excelNomer, excelGruppa, excelTovar});
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
		btnNewButton.setIcon(new ImageIcon(Tovari.class.getResource("/icon/button (5).png")));
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
		btnNewButton.setBounds(203, 306, 104, 35);
		panelFirst.add(btnNewButton);
		
		// --- Удалить --- //
		JButton btnNewButton_1 = new JButton("");
		btnNewButton_1.setIcon(new ImageIcon(Tovari.class.getResource("/icon/button (6).png")));
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
		btnNewButton_1.setBounds(317, 306, 96, 35);
		panelFirst.add(btnNewButton_1);
		
		// --- Сохранить --- //
		JButton btnNewButton_2 = new JButton("");
		btnNewButton_2.setIcon(new ImageIcon(Tovari.class.getResource("/icon/button (7).png")));
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
				XSSFSheet sheet = workbook.getSheet("Список товаров");
				if(sheet != null)   {
				    index = workbook.getSheetIndex(sheet);
				    workbook.removeSheetAt(index);
				}
				
				XSSFSheet excelSheet = workbook.createSheet("Список товаров");
				workbook.setSheetOrder("Список товаров", 1);
				
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
		btnNewButton_2.setBounds(423, 306, 112, 35);
		panelFirst.add(btnNewButton_2);
		
		///////////////////////////
		// --- Вторая панель --- //
		///////////////////////////
		JPanel panelSecond = new JPanel();
		panelSecond.setBackground(new Color(248, 248, 255));
		panelSecond.setLayout(null);
		contentPane.add(panelSecond, "2");

		// --- Label и TextField для добавления --- //
		// --- Номер товара --- //
		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setIcon(new ImageIcon(Tovari.class.getResource("/icon/button (8).png")));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBorder(null);
		lblNewLabel.setBackground(new Color(248, 248, 255));
		lblNewLabel.setBounds(202, 94, 138, 35);
		panelSecond.add(lblNewLabel);
		
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
		textField.setBounds(350, 96, 150, 30);
		panelSecond.add(textField);
		
		// --- Группа --- //
		JLabel lblNewLabel_1 = new JLabel("");
		lblNewLabel_1.setIcon(new ImageIcon(Tovari.class.getResource("/icon/button (9).png")));
		lblNewLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_1.setBorder(null);
		lblNewLabel_1.setBackground(new Color(248, 248, 255));
		lblNewLabel_1.setBounds(176, 140, 88, 35);
		panelSecond.add(lblNewLabel_1);
		
		JComboBox comboBox = new JComboBox();
		comboBox.setForeground(Color.BLACK);
		comboBox.setBorder(new LineBorder(new Color(147, 112, 219)));
		comboBox.setBackground(Color.WHITE);
		comboBox.setModel(new DefaultComboBoxModel(GruppiExcel(false)));
		comboBox.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		comboBox.setBounds(274, 142, 260, 30);
		panelSecond.add(comboBox);
		
		// --- Товар --- //
		JLabel lblNewLabel_2 = new JLabel("");
		lblNewLabel_2.setIcon(new ImageIcon(Tovari.class.getResource("/icon/button (10).png")));
		lblNewLabel_2.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_2.setBorder(null);
		lblNewLabel_2.setBackground(new Color(248, 248, 255));
		lblNewLabel_2.setBounds(198, 186, 78, 35);
		panelSecond.add(lblNewLabel_2);
		
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
		textField_1.setBounds(284, 188, 220, 30);
		panelSecond.add(textField_1);
		
		// --- Кнопка "ОК" --- //
		JButton btnNewButton_3 = new JButton("");
		btnNewButton_3.setIcon(new ImageIcon(Tovari.class.getResource("/icon/button (21).png")));
		btnNewButton_3.setFocusPainted(false);
		btnNewButton_3.setFocusable(false);
		btnNewButton_3.setBorder(null);
		btnNewButton_3.setBackground(new Color(248, 248, 255));
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 model.addRow(new Object[]{
						 textField.getText(),
						 comboBox.getSelectedItem(),
						 textField_1.getText(),
                  });
				 JOptionPane.showMessageDialog(null, "Запись добавлена успешно");
				 cl.show(contentPane, "1");
			}
		});
		
		btnNewButton_3.setBounds(360, 232, 54, 35);
		panelSecond.add(btnNewButton_3);
		
		cl.show(contentPane, "1");
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
}
