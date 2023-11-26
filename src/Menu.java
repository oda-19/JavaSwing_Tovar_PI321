import java.awt.Color;
import java.awt.EventQueue;
import java.awt.SystemColor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.JButton;
import javax.swing.ImageIcon;

public class Menu extends JFrame {

	private JPanel contentPane;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Menu frame = new Menu();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public Menu() {
		
		// --- Окно --- //
		setTitle("Меню");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 307);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		setLocationRelativeTo(null);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		// --- Убрать подсветку кнопок --- //
		UIManager.put("Button.select", SystemColor.info);
		
		// --- Кнопки --- //
		// --- Список товаров --- //
		JButton btnNewButton = new JButton("");
		btnNewButton.setIcon(new ImageIcon(Menu.class.getResource("/icon/button.png")));
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Tovari tovari = new Tovari();
				tovari.setVisible(true);
				tovari.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
			}
		});
		btnNewButton.setBorder(null);
		btnNewButton.setBackground(Color.WHITE);
		btnNewButton.setFocusPainted(false);
		btnNewButton.setFocusable(false);
		btnNewButton.setBounds(139, 25, 152, 35);
		contentPane.add(btnNewButton);
		
		// --- Список групп товаров --- //
		JButton btnNewButton_1 = new JButton("");
		btnNewButton_1.setIcon(new ImageIcon(Menu.class.getResource("/icon/button (1).png")));
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				GruppiTovarov gruppiTovarov = new GruppiTovarov();
				gruppiTovarov.setVisible(true);
				gruppiTovarov.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
			}
		});
		btnNewButton_1.setFocusPainted(false);
		btnNewButton_1.setFocusable(false);
		btnNewButton_1.setBackground(Color.WHITE);
		btnNewButton_1.setBorder(null);
		btnNewButton_1.setBounds(116, 71, 202, 35);
		contentPane.add(btnNewButton_1);
		
		// --- Список стран-импортеров --- //
		JButton btnNewButton_2 = new JButton("");
		btnNewButton_2.setIcon(new ImageIcon(Menu.class.getResource("/icon/button (2).png")));
		btnNewButton_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				StraniImporteri straniImporteri = new StraniImporteri();
				straniImporteri.setVisible(true);
				straniImporteri.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
			}
		});
		btnNewButton_2.setFocusable(false);
		btnNewButton_2.setFocusPainted(false);
		btnNewButton_2.setBackground(Color.WHITE);
		btnNewButton_2.setBorder(null);
		btnNewButton_2.setBounds(100, 117, 234, 35);
		contentPane.add(btnNewButton_2);
		
		// --- Список единиц измерения --- //
		JButton btnNewButton_3 = new JButton("");
		btnNewButton_3.setIcon(new ImageIcon(Menu.class.getResource("/icon/button (3).png")));
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				EdinisiIzmereniya edinisiIzmereniya = new EdinisiIzmereniya();
				edinisiIzmereniya.setVisible(true);
				edinisiIzmereniya.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
			}
		});
		btnNewButton_3.setFocusable(false);
		btnNewButton_3.setFocusPainted(false);
		btnNewButton_3.setBorder(null);
		btnNewButton_3.setBackground(Color.WHITE);
		btnNewButton_3.setBounds(98, 163, 238, 35);
		contentPane.add(btnNewButton_3);
		
		// --- Оперативная информация об экспорте товара --- //
		JButton btnNewButton_4 = new JButton("");
		btnNewButton_4.setIcon(new ImageIcon(Menu.class.getResource("/icon/button (4).png")));
		btnNewButton_4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				OperativInform operativInform = new OperativInform();
				operativInform.setVisible(true);
				operativInform.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
			}
		});
		btnNewButton_4.setFocusable(false);
		btnNewButton_4.setFocusPainted(false);
		btnNewButton_4.setBorder(null);
		btnNewButton_4.setBackground(Color.WHITE);
		btnNewButton_4.setBounds(25, 209, 387, 35);
		contentPane.add(btnNewButton_4);
	}
}
