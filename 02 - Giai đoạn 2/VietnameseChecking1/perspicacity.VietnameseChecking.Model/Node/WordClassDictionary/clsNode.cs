// Dự án:   BẮT LỖI CHÍNH TẢ TIẾNG VIỆT
// Tác giả: Lê Viết Mẫn
// Email:   manleviet@yahoo.com.hk

using System;

namespace perspicacity.VietnameseChecking.Model.Node.WordClassDictionary
{
	/// <summary>
	/// Lớp _node này biểu diễn cho một nút trong cấu trúc
	/// từ điển WordClass.
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
		/// <param name="wordclass">
		/// Nhận chuỗi từ loại.
		/// </param>
		/// <param name="sense">
		/// Nhận chuỗi ý nghĩa của từ loại.
		/// </param>
		public _node(string wordclass, string sense) : base(wordclass, sense){}
		#endregion

		#region Vùng thuộc tính
		/// <summary>
		/// Thuộc tính WordClass, lấy ra hay gán từ loại.
		/// </summary>
		public string WordClass
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
		/// Thuộc tính Sense, lấy ra hay gán ý nghĩa từ loại.
		/// </summary>
		public string Sense
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
