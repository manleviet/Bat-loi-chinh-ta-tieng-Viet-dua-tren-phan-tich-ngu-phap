// Dự án:   BẮT LỖI CHÍNH TẢ TIẾNG VIỆT
// Tác giả: Lê Viết Mẫn
// Email:   manleviet@yahoo.com.hk

using System;

namespace perspicacity.VietnameseChecking.Model.Node.WordDictionary
{
	/// <summary>
	/// Lớp _node này biểu diễn cho một nút trong cấu trúc
	/// từ điển Word.
	/// </summary>
	public class _node : _base
	{
		#region Vùng hàm khởi tạo
		/// <summary>
		/// Cấu tử cơ bản, không có tham số.
		/// </summary>
		public _node():base(){}

		/// <summary>
		/// Cấu tử thứ hai, có hai tham số.
		/// </summary>
		/// <param name="word">
		/// Nhận chuỗi từ.
		/// </param>
		/// <param name="wordclass">
		/// Nhận chuỗi từ loại của từ.
		/// </param>
		public _node(string word, string wordclass) : base(word, wordclass){}
		#endregion

		#region Vùng thuộc tính
		/// <summary>
		/// Thuộc tính Word, lấy ra hay gán một từ.
		/// </summary>
		public string Word
		{
			get
			{
				return base.Content1;
			}
			set
			{
				base.Content1 = value;
			}
		}

		/// <summary>
		/// Thuộc tính WordClass, lấy ra hay gán chuỗi từ loại của từ.
		/// </summary>
		public string WordClass
		{
			get
			{
				return base.Content2;
			}
			set
			{
				base.Content2 = value;
			}
		}
		#endregion
	}
}