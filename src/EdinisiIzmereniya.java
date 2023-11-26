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

public class EdinisiIzmereniya extends JFrame {

	private JPanel contentPane;
	private static JTable table;
	private JTextField textField;
	private JScrollPane scrollPane;
    
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					EdinisiIzmereniya frame = new EdinisiIzmereniya();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public EdinisiIzmereniya() {	
		
		// --- Окно --- //
		setForeground(SystemColor.control);
		setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
		setTitle("Список единиц измерения");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(0, 0, 622, 406);
		setLocationRelativeTo(null);
		contentPane = new JPanel();
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
		panelFirst.setBackground(new Color(248, 248, 255));
		panelFirst.setLayout(null);
		contentPane.add(panelFirst, "1");
		
		// --- Таблица --- //
		scrollPane = new JScrollPane();
		scrollPane.setBorder(null);
		scrollPane.setForeground(Color.BLACK);
		scrollPane.setBackground(Color.WHITE);
		scrollPane.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		scrollPane.setBounds(25, 25, 553, 270);
		
		table = new JTable(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"Код товара", "Товар", "Единица измерения"
			}
		));
		table.setForeground(Color.BLACK);
		table.setBackground(Color.WHITE);
		table.setSelectionBackground(new Color(216, 191, 216));
		table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		Font bigFont = new Font("Comic Sans MS", Font.PLAIN, 16);
		table.getTableHeader().setFont(bigFont);
		table.getTableHeader().setBackground(Color.WHITE);
		
		table.getColumnModel().getColumn(0).setPreferredWidth(105);
		table.getColumnModel().getColumn(1).setPreferredWidth(270);
		table.getColumnModel().getColumn(2).setPreferredWidth(160);

		table.setRowHeight(30);
		
		scrollPane.setViewportView(table);
		table.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		panelFirst.add(scrollPane);
		
		DefaultTableModel model = (DefaultTableModel)table.getModel();
		
		// --- Добавление в таблицу Combobox с данными из Excel --- //
		TovariExcel(true);
		
		// --- Добавление в таблицу Combobox --- //
		JComboBox comboBox = new JComboBox();
		comboBox.setEditable(true);
		comboBox.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		comboBox.setModel(new DefaultComboBoxModel(new String[] {"тонна (т)", "тысяча (тыс.)", "куб. м"}));
		table.getColumnModel().getColumn(2).setCellEditor(new DefaultCellEditor(comboBox));
		
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
				XSSFSheet excelSheet = workbook.getSheetAt(4);
				
				try {
					for (int row = 1; row <= excelSheet.getLastRowNum(); row++) {
						XSSFRow excelRow = excelSheet.getRow(row);

		                XSSFCell excelKodTovara = excelRow.getCell(0);
		                XSSFCell excelTovar = excelRow.getCell(1);
		                XSSFCell excelEdinisaIzmereniya = excelRow.getCell(2);
		                    
		                model.addRow(new Object[]{excelKodTovara, excelTovar, excelEdinisaIzmereniya});
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
		btnNewButton.setIcon(new ImageIcon(EdinisiIzmereniya.class.getResource("/icon/button (5).png")));
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
		btnNewButton.setBounds(135, 306, 104, 35);
		panelFirst.add(btnNewButton);
		
		// --- Удалить --- //
		JButton btnNewButton_1 = new JButton("");
		btnNewButton_1.setIcon(new ImageIcon(EdinisiIzmereniya.class.getResource("/icon/button (6).png")));
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
		btnNewButton_1.setBounds(249, 306, 96, 35);
		panelFirst.add(btnNewButton_1);
		
		// --- Сохранить --- //
		JButton btnNewButton_2 = new JButton("");
		btnNewButton_2.setIcon(new ImageIcon(EdinisiIzmereniya.class.getResource("/icon/button (7).png")));
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
				XSSFSheet sheet = workbook.getSheet("Список единиц измерения");
				if(sheet != null)   {
				    index = workbook.getSheetIndex(sheet);
				    workbook.removeSheetAt(index);
				}
				
				XSSFSheet excelSheet = workbook.createSheet("Список единиц измерения");
				workbook.setSheetOrder("Список единиц измерения", 4);
				
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
		btnNewButton_2.setBounds(355, 306, 112, 35);
		panelFirst.add(btnNewButton_2);
		
		///////////////////////////
		// --- Вторая панель --- //
		///////////////////////////
		JPanel panelSecond = new JPanel();
		panelSecond.setBackground(new Color(248, 248, 255));
		panelSecond.setLayout(null);
		contentPane.add(panelSecond, "2");

		// --- Label и TextField для добавления --- //
		// --- Код товара --- //
		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setIcon(new ImageIcon(EdinisiIzmereniya.class.getResource("/icon/button (14).png")));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBorder(null);
		lblNewLabel.setBackground(new Color(248, 248, 255));
		lblNewLabel.setBounds(166, 96, 116, 35);
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
		textField.setBounds(292, 98, 150, 30);
		panelSecond.add(textField);
		
		// --- Товар --- //
		JLabel lblNewLabel_1 = new JLabel("");
		lblNewLabel_1.setIcon(new ImageIcon(EdinisiIzmereniya.class.getResource("/icon/button (10).png")));
		lblNewLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_1.setBorder(null);
		lblNewLabel_1.setBackground(new Color(248, 248, 255));
		lblNewLabel_1.setBounds(132, 142, 78, 35);
		panelSecond.add(lblNewLabel_1);
		
		JComboBox comboBox_1 = new JComboBox();
		comboBox_1.setForeground(Color.BLACK);
		comboBox_1.setBorder(new LineBorder(new Color(147, 112, 219)));
		comboBox_1.setBackground(Color.WHITE);
		comboBox_1.setModel(new DefaultComboBoxModel(TovariExcel(false)));
		comboBox_1.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		comboBox_1.setBounds(220, 144, 260, 30);
		panelSecond.add(comboBox_1);
		
		// --- Единица измерения --- //
		JLabel lblNewLabel_2 = new JLabel("");
		lblNewLabel_2.setIcon(new ImageIcon(EdinisiIzmereniya.class.getResource("/icon/button (15).png")));
		lblNewLabel_2.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_2.setBorder(null);
		lblNewLabel_2.setBackground(new Color(248, 248, 255));
		lblNewLabel_2.setBounds(142, 185, 186, 35);
		panelSecond.add(lblNewLabel_2);
		
		JComboBox comboBox_2 = new JComboBox();
		comboBox_2.setForeground(Color.BLACK);
		comboBox_2.setBorder(new LineBorder(new Color(147, 112, 219)));
		comboBox_2.setBackground(Color.WHITE);
		comboBox_2.setModel(new DefaultComboBoxModel(new String[] {"тонна (т)", "тысяча (тыс.)", "куб. м"}));
		comboBox_2.setFont(new Font("Comic Sans MS", Font.PLAIN, 16));
		comboBox_2.setBounds(338, 187, 130, 30);
		panelSecond.add(comboBox_2);
		
		// --- Кнопка "ОК" --- //
		JButton btnNewButton_3 = new JButton("");
		btnNewButton_3.setIcon(new ImageIcon(EdinisiIzmereniya.class.getResource("/icon/button (21).png")));
		btnNewButton_3.setFocusPainted(false);
		btnNewButton_3.setFocusable(false);
		btnNewButton_3.setBorder(null);
		btnNewButton_3.setBackground(new Color(248, 248, 255));
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 model.addRow(new Object[]{
						 textField.getText(),
						 comboBox_1.getSelectedItem(),
						 comboBox_2.getSelectedItem(),
                  });
				 JOptionPane.showMessageDialog(null, "Запись добавлена успешно");
				 cl.show(contentPane, "1");
			}
		});
		
		btnNewButton_3.setBounds(274, 231, 54, 35);
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
