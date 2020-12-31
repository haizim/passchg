push(){
	git add .
	echo "------------added----------"
	echo "masukkan commit :"
	read commit
	git commit -m "$commit"
	echo "------------committed----------"
	git push -u origin master 
	echo "------------pushed----------"
}

echo "tambahkan baru ?y/n"
read tambah
if [[ $tambah = "y" ]];
then
	echo "masukkan alamat repo (https://github.com/mmm/www.git)"
	read repo
	git init
	git remote add origin $repo
	push
elif [[ $tambah = "n" ]];
then
	push
else
	echo "input tidak cocok"
fi
