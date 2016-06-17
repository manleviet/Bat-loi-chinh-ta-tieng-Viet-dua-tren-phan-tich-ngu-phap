using System;
using System.Collections;

namespace LinkedList
{
	/// <summary>
	/// Một danh sách liên kết đôi cài đặt giao diện <see cref="IList"/>.
	/// </summary>
	/// <remarks>
	/// <see cref="LinkedList"/> có thể sử dụng cho tất cả mọi đối tượng kể cả null.
	/// <para>Lớp này chưa đảm bảo về an toàn tiểu trình.</para>
	/// Phương thức <see cref="IEnumerator.MoveNext()"/> của <see cref="IEnumerator"/> mà nó đựơc trả
	/// ra từ <see cref="LinkedList.GetEnumerator"/> là dễ gặp lỗi. Nếu <see cref="LinkedList"/> được
	/// sửa đổi sau khi enumerator được tạo ra, nó sẽ ném ra một ngoại lệ <see cref="SystemException"/>.
	/// Hành vi này là không an toàn, và không nên phụ thuộc vào nó.
	/// </remarks>
	[Serializable]
	public class LinkedList : IList, ICloneable
	{
		#region Khai báo biến
		private Node headerNode;
		private int  count;
		private int  modifications;
		#endregion

		#region Hàm khởi tạo
		/// <summary>
		/// Tạo ra một thể hiện mới của lớp LinkedList và nó là rỗng.
		/// </summary>
		public LinkedList()
		{
			headerNode = new Node(null, null, null);
			headerNode.NextNode = headerNode;
			headerNode.PreviousNode = headerNode;
			count = 0;
		}

		public LinkedList(ICollection collection):this()
		{
			AddAll(collection);
		}
		#endregion
		
		#region Thuộc tính
		public virtual int Count
		{
			get
			{
				return count;
			}
		}
		
		public virtual bool IsSynchronized
		{
			get
			{
				return false;
			}
		}

		public virtual object SyncRoot
		{
			get
			{
				return this;
			}
		}

		public virtual bool IsFixedSize
		{
			get
			{
				return false;
			}
		}

		public virtual bool IsReadOnly
		{
			get
			{
				return false;
			}
		}

		public virtual object this[int index]
		{
			get
			{
				return FindNodeAt(index).CurrentNode;
			}
			set
			{
				FindNodeAt(index).CurrentNode = value;
			}
		}
		#endregion

		public int Add(object value)
		{
			Insert(count, value);
			return (count - 1);
		}

		public void AddAll(ICollection collection)
		{
			InsertAll(count, collection);
		}

		public void Clear()
		{
			modifications++;
			headerNode.NextNode = headerNode;
			headerNode.PreviousNode = headerNode;
			count = 0;
		}

		public object Clone()
		{
			LinkedList listClone = new LinkedList();
			for (Node node = headerNode.NextNode; node != headerNode; node = node.NextNode)
				listClone.Add(node.CurrentNode);
			return listClone;
		}

		public LinkedList Clone(bool attemptDeepCopy)
		{
			LinkedList listClone;
			if (attemptDeepCopy)
			{
				listClone = new LinkedList();
				object currentObject;
				for (Node node = headerNode.NextNode; node != headerNode; node = node.NextNode)
				{
					currentObject = node.CurrentNode;
					if (currentObject == null)
						listClone.Add(null);
					else if (currentObject is ICloneable)
						listClone.Add(((ICloneable)currentObject).Clone());
					else
						throw new SystemException("The object of type [" + currentObject.GetType() +
											"] in the list is not an ICloneable, cannot attempt a deep copy.");
                }
			}
			else
				listClone = (LinkedList)this.Clone();
			return listClone;
		}

		public bool Contains(object value)
		{
			return (0 <= IndexOf(value));
		}

		public void CopyTo(Array array, int index)
		{
			if (array != null)
			{
				if (0 <= index)
				{
					if (array.Rank == 1)
					{
						if (count <= (array.Length - index))
						{
							if (index < array.Length)
							{
								for (int i = index, j = 0; j < count; i++, j++)
									array.SetValue(FindNodeAt(j).CurrentNode, i);
							}
							else
								throw new ArgumentException("index is equal to or greater than the length of array.", "index");
						}
						else
							throw new ArgumentException("The number of elements is greater than the available space from index in the destination array.", "array");
					}
					else
						throw new ArgumentException("Multidimensional array", "array");
				}
				else
					throw new ArgumentOutOfRangeException("index", index, "less than zero");
			}
			else
				throw new ArgumentNullException("array");
		}

		public IEnumerator GetEnumerator()
		{
			return new LinkedListEnumerator(this);
		}

		public int IndexOf(object value)
		{
			int currentIndex = 0;
			if (value == null)
			{
				for (Node node = headerNode.NextNode; node != headerNode; node = node.NextNode)
				{
					if (node.CurrentNode == null)
						break;
					currentIndex++;
				}
			}
			else
			{
				for (Node node = headerNode; node != headerNode; node = node.NextNode)
				{
					if (value.Equals(node.CurrentNode))
						break;
					currentIndex++;
				}
			}
			if (count <= currentIndex)
				currentIndex = -1;
			return currentIndex;
		}

		public void Insert(int index, object value)
		{
			Node node;
			if (index == count)
				node = new Node(value, headerNode, headerNode.PreviousNode);
			else
			{
				Node tmp = FindNodeAt(index);
				node = new Node(value, tmp, tmp.PreviousNode);
			}
			node.PreviousNode.NextNode = node;
			node.NextNode.PreviousNode = node;
			count++;
			modifications++;
		}

		public void InsertAll(int index, ICollection collection)
		{
			if (collection != null)
			{
				if (0 <= collection.Count)
				{
					modifications++;
					Node startingNode = (index == count ? headerNode : FindNodeAt(index));
					Node previousNode = (startingNode.PreviousNode);
					foreach (object obj in collection)
					{
						Node node = new Node(obj, startingNode, previousNode);
						previousNode.NextNode = node;
						previousNode = node;
					}
					startingNode.PreviousNode = previousNode;
					count += collection.Count;
				}
				else
					throw new ArgumentOutOfRangeException("index", index, "less than zero");
			}
			else
				throw new ArgumentNullException("collection");
		}

		public void Remove(object value)
		{
			if (value == null)
			{
				for (Node node = headerNode.NextNode; node != headerNode; node = node.NextNode)
					if (node.CurrentNode == null)
						Remove(node);
			}
			else
			{
				for (Node node = headerNode; node != headerNode; node = node.NextNode)
					if (value.Equals(node.CurrentNode))
						Remove(node);
			}
		}

		public void RemoveAt(int index)
		{
			Remove(FindNodeAt(index));
		}

		#region Hàm private
		private Node FindNodeAt(int index)
		{
			if (index < 0 || count <= index)
				throw new IndexOutOfRangeException("Attempted to access index " + index + 
							", while the total coutn is " + count + ".");
			Node node = headerNode;
			if (index < (count/2))
			{
				for (int i = 0; i <= index; i++)
					node = node.NextNode;
			}
			else
			{
				for (int i = count; i > index; i--)
					node = node.PreviousNode;
			}
			return node;
		}

		private void Remove(Node value)
		{
			if (value != headerNode)
			{
				value.PreviousNode.NextNode = value.NextNode;
				value.NextNode.PreviousNode = value.PreviousNode;
				count--;
				modifications++;
			}
		}
		#endregion

		[Serializable]
		private class Node
		{
			#region Khai báo biến
			private object currentNode;
			private Node nextNode;
			private Node previousNode;
			#endregion

			#region Hàm khởi tạo
			public Node(object currentNode, Node nextNode, Node previousNode)
			{
				this.currentNode = currentNode;
				this.nextNode = nextNode;
				this.previousNode = previousNode;
			}
			#endregion

			#region Thuộc tính
			public object CurrentNode
			{
				get
				{
					return currentNode;
				}
				set
				{
					currentNode = value;
				}
			}

			public Node NextNode
			{
				get
				{
					return nextNode;
				}
				set
				{
					nextNode = value;
				}
			}

			public Node PreviousNode
			{
				get
				{
					return previousNode;
				}
				set
				{
					previousNode = value;
				}
			}
			#endregion
		}

		[Serializable]
		private class LinkedListEnumerator: IEnumerator
		{
			#region Khai báo biến
			private LinkedList linkedList;
			private int validModificationCount;
			private Node currentNode;
			#endregion

			public LinkedListEnumerator(LinkedList linkedlist)
			{
				this.linkedList = linkedlist;
				validModificationCount = linkedlist.modifications;
				currentNode = linkedlist.headerNode;
			}

			public object Current
			{
				get
				{
					return currentNode.CurrentNode;
				}
			}

			public void Reset()
			{
				currentNode = linkedList.headerNode;
			}

			public bool MoveNext()
			{
				bool moveSuccessful = false;
				if (validModificationCount != linkedList.modifications)
					throw new SystemException("A concurrent modification occured to the LinkedList while accessing it through it's enumerator.");
				currentNode = currentNode.NextNode;
				if (currentNode != linkedList.headerNode)
					moveSuccessful = true;
				return moveSuccessful;
			}
		}
	}
}
