   M S B A M L   `   `   `  ����  GetCountryCodes/    )GetCountryCodes.frmCountryCarriersPreview   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$!  Country & Carrier Get Names��5   	   $	��366��61   $	��525��6>    � 
ResizeMode$ NoResize=�6J    � Icon$7 1/GetCountryCodes;component/Images/fregat-oleg.jpg��6`    � WindowStartupLocation$ CenterScreen=�6�    � WindowStyle$ SingleBorderWindow=�6�   ���	#0062454E.��5      � ���	#FFB4B4B4.+�5   
   �� -   5        $��label1��5   o   $!��_Select columns to be used:��6   $��28��66   $
��Left=�6B   $��13,8,0,0q�6]   $	��Top=�6}   $	��161��6�   &   System.Windows.Controls.DataGrid  5   
   -   5        $��gridExcelPreview��5   i      AutoGenerateColumns . $	��230��6.   $
��Left=�6;   $��	12,42,0,0q�6V   $	��Top=�6�   $	��479��6�      CanUserReorderColumns .     CanUserResizeColumns .   � DataContext 6�   �� 6�   �� 5   
   -   5        $��pbGetCountryCarriers��5   x   $"��_Get country & carrier names��6   $��31��6<   $��Right=�6H   $��
0,285,12,0q�6d   $	��Top=�6�   $	��153��6�   6�   �� 5	   
   -   5        $��pbBrowse��5	   r   $ ��_Browse file with codes...��6   $��31��66   $
��Left=�6B   $��179,285,0,0q�6]   $	��Top=�6�   $	��153��6�   6�   5
      5      