import java.awt.Point;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Stack;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class BoardFrame {
	
	private JFileChooser saver = new JFileChooser() ;
	
	private Stack<Point> stack = new Stack<Point>() ;
	private int[][] boardPlayed ;
	private ArrayList<int[][]> steps = new ArrayList<int[][]>() ;
	
	private int DIM ;
	
	SXSSFWorkbook workbook ;
	CellStyle board ;

	// Letter->粂ēダM->籔じ皌才腹(竊腹单单)Z->フだ筳腹S->才腹N->NumberP->夹翴才腹C->ㄤ(北じ单)
	Pattern special = Pattern.compile("\\p{P}|\\p{Z}|\\p{L}|\\p{S}|\\p{M}|\\p{C}") ;
	Matcher finder ;
	
	public BoardFrame() {
		
		ask();
		createXlsx() ;
		checkmate() ;
		createContent() ;
		save() ;
		
	}
	
	public void createXlsx() {
		workbook = new SXSSFWorkbook() ;
		board = workbook.createCellStyle() ;
		board.setBorderLeft( BorderStyle.THIN ) ;
		board.setBorderBottom( BorderStyle.THIN ) ;
		board.setBorderRight( BorderStyle.THIN ) ;
		board.setBorderTop( BorderStyle.THIN ) ;
		board.setLeftBorderColor( IndexedColors.BLACK.getIndex() ) ;
		board.setBottomBorderColor( IndexedColors.BLACK.getIndex() ) ;
		board.setRightBorderColor( IndexedColors.BLACK.getIndex() ) ;
		board.setTopBorderColor( IndexedColors.BLACK.getIndex() ) ;
	}
	
	public void ask() {
		
		String input = JOptionPane.showInputDialog(null, "Please enter board size.", "Eight Queen", JOptionPane.DEFAULT_OPTION) ;
		
		if( input == null ) {
			System.exit(0) ;
		}
		else {
			finder = special.matcher( input ) ;
			
			if( finder.find() || input.equals("") || input.equals("0") ) {
				JOptionPane.showMessageDialog(null, "Wrong input ! Please enter again.", "Wrong Input", JOptionPane.DEFAULT_OPTION) ;
				ask() ;
			}
			else {
				DIM = Integer.parseInt( input ) ;
				boardPlayed = new int[DIM][DIM] ;
			}
			
		}
	}
	
	public void checkmate() {
		
		int i = 1 ;
			
		int j = 0 ;
		
		while( i <= DIM ) {
			
			j++ ;

			if( !( boardPlayed[(i-1)][j-1] == -1 )  ) {
				boardPlayed[(i-1)][j-1] = 1 ;
				stack.push( new Point( j, i ) ) ;
				int[][] transBoard = new int[DIM][DIM] ;
				for( int row = 0 ; row < DIM ; row++ ) {
					for( int col = 0 ; col < DIM ; col++ ) {
						transBoard[row][col] = boardPlayed[row][col] ;
					}
				}
				steps.add( transBoard ) ;
				
				for( int row = 1 ; row <= DIM ; row++ ) {
					for( int col = 1 ; col <= DIM ; col++ ) {
						
						if( !(row == i && col == j) ) {
							if( ( col == j || col-j == i-row || col-j == row-i ) && i < row ) {
								
								boardPlayed[row-1][col-1] = -1 ;

							}
						}
						
					}
				}
				
				i++ ;
				j = 0 ;
			}

			while( j >= DIM ) {

				Point removedPoint = stack.pop() ;
				boardPlayed[removedPoint.y-1][removedPoint.x-1] = 0 ;
				
				int[][] transBoard = new int[DIM][DIM] ;
				for( int row = 0 ; row < DIM ; row++ ) {
					for( int col = 0 ; col < DIM ; col++ ) {
						transBoard[row][col] = boardPlayed[row][col] ;
					}
				}
				steps.add( transBoard ) ;
				
				i = removedPoint.y ;
				j = removedPoint.x ;
				
				for( int row = 1 ; row <= DIM ; row++ ) {
					for( int col = 1 ; col <= DIM ; col++ ) {
						
						if( !(row == i && col == j) ) {
							if( ( col == j || col-j == i-row || col-j == row-i ) && i < row ) {
								boardPlayed[row-1][col-1] = 0 ;
							}
						}
						
					}
				}
				
				stack.stream().forEach( point -> {
					int c = point.x ;
					int r = point.y ;
					
					for( int row = 1 ; row <= DIM ; row++ ) {
						for( int col = 1 ; col <= DIM ; col++ ) {
							if( !(row == r && col == c) ) {
								if( ( col == c || col-c == r-row || col-c == row-r ) && r-row < 0 ) {
									boardPlayed[row-1][col-1] = -1 ;
								}
							}
						}
					}
					
				} ) ;
				
			}
		}
		
	}
	
	public void createContent() {
		
		Iterator<int[][]> step = steps.iterator() ;
		int pageNum = 1 ;
		int serialNum = 1 ;
		int rowIndex = 0 ;
		int columnIndex = 0 ;
		Boolean hasRow = false ;
		Sheet sheet = workbook.createSheet( new StringBuffer().append("Page. ").append( String.valueOf(pageNum) ).toString() ) ;
		sheet.setDefaultColumnWidth( 2 ) ;
		while( step.hasNext() ) {
			if( serialNum % 1000 == 0 ) {
				hasRow = false ;
				pageNum ++ ;
				rowIndex = 0 ;
				columnIndex = 0 ;
				sheet = workbook.createSheet( new StringBuffer().append("Page. ").append( String.valueOf(pageNum) ).toString() ) ;
				sheet.setDefaultColumnWidth( 2 ) ;
			}
			if( !hasRow ) {
				sheet.createRow( rowIndex ) ;
			}
			sheet.getRow( rowIndex ).createCell( columnIndex ).setCellValue( String.valueOf( serialNum ) ) ;
			columnIndex ++ ;
			for( int[] row : step.next() ) {
				rowIndex ++ ;
				if( !hasRow ) {
					sheet.createRow( rowIndex ) ;
				}
				for( int cell : row ) {
					if( cell == 0 || cell == -1 ) {
						sheet.getRow( rowIndex ).createCell( columnIndex ).setCellStyle( board ) ;
						columnIndex ++ ;
					}
					else {
						sheet.getRow( rowIndex ).createCell( columnIndex ).setCellValue("O") ;
						sheet.getRow( rowIndex ).getCell( columnIndex ).setCellStyle( board ) ;
						columnIndex ++ ;
					}
				}
				if( !( (rowIndex+1) % (DIM+1) == 0 ) && columnIndex % (DIM+1) == 0  ) {
					columnIndex = columnIndex - DIM ;
				}
				if( (rowIndex+1) % (DIM+1) == 0 && columnIndex == DIM+1 ) {
					hasRow = true ;
				}
			}
			serialNum ++ ;
			if( ( columnIndex + DIM+1 ) > 58 ) {
				rowIndex ++ ;
				columnIndex = 0 ;
				hasRow = false ;
			}
			else {
				rowIndex = rowIndex - DIM ;
			}
			
		}
		
	}
	
	public void save() {
		
		saver.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY) ;
		saver.showSaveDialog(null) ;
		try {
			workbook.write( Files.newOutputStream( Paths.get( new StringBuffer()
					.append( saver.getSelectedFile().getAbsolutePath() )
					.append("\\8Queen_")
					.append(DIM)
					.append("_")
					.append(DIM)
					.append(".xlsx").toString() ) ) ) ;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
}
